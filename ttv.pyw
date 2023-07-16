import tkinter as tk
from calendar import day_name
from os import mkdir
from tkinter import ttk, messagebox, Toplevel
from tkinter.filedialog import askopenfilename

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import image


# Time Table Visualizer
# Kevin Yu


class Application(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("TTV")
        self.root.resizable(False, False)

        self.filepath_label = ttk.Label(self.root, text="Open Spreadsheet file", width=25, anchor='w')
        self.imagepath_label = ttk.Label(self.root, text="Open logo image", width=25, anchor='w')
        self.filepath_button = ttk.Button(self.root, text="Open file", command=self.open_spreadsheet, width=11)
        self.imagepath_button = ttk.Button(self.root, text="Open Image", command=self.open_image, width=11)
        self.generate_button = ttk.Button(self.root, text="Generate!", command=self.generate, width=11)
        self.exit_button = ttk.Button(self.root, text="Exit", command=self.root.destroy)

        self.filepath_label.grid(row=0, column=0, sticky='W')
        self.filepath_button.grid(row=0, column=1, sticky='W')
        self.imagepath_label.grid(row=1, column=0, sticky='W')
        self.imagepath_button.grid(row=1, column=1, sticky='W')
        self.generate_button.grid(row=2, column=1)
        self.exit_button.grid(row=2, column=0, sticky='W')

        self.root.mainloop()

    def open_spreadsheet(self):
        filetypes = (('Excel Spreadsheet', '*.xlsx'), ('All files', '*.*'))
        file_path = askopenfilename(filetypes=filetypes, initialdir=".\\", title="Locate timetable spreadsheet")
        if file_path:
            self.filepath_label.config(text=file_path)

    def open_image(self):
        filetypes = (('PNG Image', '*.png'), ('All files', '*.*'))
        image_path = askopenfilename(filetypes=filetypes, initialdir=".\\", title="Locate logo image")
        if image_path:
            self.imagepath_label.config(text=image_path)

    def generate(self):
        excel_path = self.filepath_label.cget('text')
        image_path = self.imagepath_label.cget('text')
        client = TimeTableVis(self, excel_path, image_path)

    def error(self):
        messagebox.showerror('Error', "Error: Failed Excel spreadsheet comprehension!\n"
                                      "Try removing any active filters in the file, and make sure the data starts at "
                                      "cell A1, so the first row is the column headings and the second row is the "
                                      "first event/class/entry.")
        self.child.destroy()

    def prog_window(self):
        self.child = Toplevel(self.root)
        self.child.geometry("300x120")
        self.child.resizable(False, False)
        self.child.title("TTV - Generating Images...")
        self.progbar = ttk.Progressbar(self.child, orient='horizontal', mode='determinate', length=200)
        self.prog_label = ttk.Label(self.child, text="Initialising...", width=45, anchor='w')
        self.detail_label = ttk.Label(self.child, text="Initialising...", width=45, anchor='w')
        self.progbar.grid(row=0, column=0, columnspan=2, padx=10, pady=20)
        self.prog_label.grid(row=1, column=0, columnspan=2, padx=10)
        self.detail_label.grid(row=2, column=0, columnspan=2, padx=10)
        self.root.update()


class TimeTableVis(object):
    def __init__(self, tkobj, excel_path, image_path):
        self.excel_path = excel_path
        self.timetable = {}
        self.timeframe = ""  # year and term number
        self.tkobj = tkobj
        save_path = excel_path.split("/")
        save_path = "/".join(save_path[0:-1])

        tkobj.prog_window()
        self.comprehend_excel_file()
        self.generate_images(save_path, image_path)
        messagebox.showinfo('Complete', 'Timetable generation has been completed.\nIt is located in the same directory'
                                        ' as the spreadsheet file, under the folder "output".')
        self.tkobj.root.destroy()

    def comprehend_excel_file(self):
        try:
            file = pd.read_excel(self.excel_path)
        except ValueError as e:
            self.tkobj.error()
            return
        self.tkobj.progbar['value'] = 0
        self.tkobj.prog_label.config(text="Comprehending excel sheet file...")
        self.timeframe = f"{file['Year'][0]} - {file['Study Period - Code'][0]}"
        for row in file.iloc():
            date = row["Start Date"]
            dayoftheweek = day_name[date.weekday()]
            if row["Building Id"] not in self.timetable.keys():
                self.timetable[row["Building Id"]] = {'Monday': {}, 'Tuesday': {}, 'Wednesday': {},
                                                      'Thursday': {}, 'Friday': {}, 'Saturday': {},
                                                      'Sunday': {}}
                self.timetable[row["Building Id"]][dayoftheweek][row["Room Id"]] = [row]
            elif row["Room Id"] not in self.timetable[row["Building Id"]][dayoftheweek].keys():
                self.timetable[row["Building Id"]][dayoftheweek][row["Room Id"]] = [row]
            else:
                self.timetable[row["Building Id"]][dayoftheweek][row["Room Id"]].append(row)
            self.tkobj.progbar['value'] += 100 / (len(file.index) - 1)
            self.tkobj.root.update()
        self.standardize_rooms_and_levels()

    def standardize_rooms_and_levels(self):
        """
        This function will make sure that all campuses are also split apart by their levels, and
        make sure every classroom in the campus is present in the timetable.

        Make sure that the third last character (number or letter) denotes the level of the classrooms, such as:
        OC206 --> O'Connell Street level 2, room 6
        M104 --> Market Street level 1, room 4
        NCG02 --> New Castle level G (ground), room 2
        Additionally, if the class is online, keep the first 3 letters of the campus as ONL, otherwise it will split
        the online timetable into different folders.
        """
        copytable = self.timetable.copy()
        self.tkobj.prog_label.config(text="Standardizing rooms and levels...")
        self.tkobj.progbar['value'] = 0
        for campus in copytable:
            rooms = []
            for dayofweek in self.timetable[campus]:
                for room in self.timetable[campus][dayofweek]:
                    rooms.append(room)
            for dayofweek in self.timetable[campus]:
                for room in rooms:
                    if room not in self.timetable[campus][dayofweek].keys():
                        self.timetable[campus][dayofweek][room] = []
            if campus[0:3] == 'ONL':  # don't bother splitting into levels if it's online
                continue
            levels = set()
            for room in rooms:
                levels.add(str(room)[-3])
            if len(levels) > 1:
                self.tkobj.detail_label.config(text=f"Multiple levels found on campus {campus}")
                for level in levels:
                    self.tkobj.detail_label.config(text=f"Adding classes for {campus} Level {level}")
                    self.timetable[f'{campus} Level {level}'] = {'Monday': {}, 'Tuesday': {}, 'Wednesday': {},
                                                                 'Thursday': {}, 'Friday': {}, 'Saturday': {},
                                                                 'Sunday': {}}
                    for room in rooms:
                        if str(room)[-3] == level:
                            for dayofweek in self.timetable[f'{campus} Level {level}']:
                                self.timetable[f'{campus} Level {level}'][dayofweek][room] = [cls for cls in
                                                                                              self.timetable[campus][
                                                                                                  dayofweek][room] if
                                                                                              cls['Room Id'][
                                                                                                  -3] == level]

                del self.timetable[campus]
            self.tkobj.progbar['value'] += 100 / (len(copytable.keys()) - 1)
            self.tkobj.root.update()

    def text_display_timetable(self):
        """
        Displays the timetable in text format just for testing
        """
        for campus in self.timetable:
            print(campus)
            for dayofweek in self.timetable[campus]:
                print("--" + dayofweek)
                for room in self.timetable[campus][dayofweek]:
                    print("----" + room)
                    for cls in self.timetable[campus][dayofweek][room]:
                        print("------", cls["Start Time"])

    def generate_images(self, save_path, image_path):
        """
        Actually does the generation of timetable images
        """
        self.tkobj.prog_label.config(text="Generating timetables...")
        self.tkobj.detail_label.config(text="Initialising...")
        self.tkobj.progbar['value'] = 0
        self.tkobj.root.update()
        for campus in self.timetable:
            if campus[0:3] == 'ONL':
                # don't draw online timetable
                continue
            for dayofweek in self.timetable[campus]:
                self.tkobj.detail_label.config(text=f"Plotting campus {campus} for {dayofweek}")
                fig, ax = plt.subplots(figsize=(12, 7))
                rooms = []
                for room in self.timetable[campus][dayofweek]:
                    rooms.append(room)
                rooms.sort(reverse=True)

                # Axis formatting
                ax.xaxis.grid(zorder=0)  # zorder is to hide it behind bars
                ax.set_ylim(0.3, len(rooms) + 0.7)
                ax.set_xlim(8.9, 21.9)
                ax.set_yticks(range(1, len(rooms) + 1))
                ax.set_yticklabels(rooms)
                ax.set_xticks(range(9, 23))
                ax.set_xticklabels(["09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00",
                                    "18:00", "19:00", "20:00", "21:00", "22:00"])
                ax.set_xlabel('Time')

                # Adding logo - logo name must be logo.png in the same directory
                logo = image.imread(image_path)
                fig.figimage(logo, 10, 10)

                # Plotting Classes
                for roomindex in range(len(rooms)):
                    if roomindex != 0:
                        # line in between boxes/ygrid
                        ax.axline((1, roomindex + 0.5), (2, roomindex + 0.5), color='grey', linewidth=0.5)
                    for cls in self.timetable[campus][dayofweek][rooms[roomindex]]:
                        start = float(cls["Start Time"]) / 10000
                        end = float(cls["End Time"]) / 10000

                        bar = plt.bar(x=start, height=0.8, width=end - start - 0.05, bottom=roomindex + 0.6,
                                      color='#d9d9d9', ec='black', zorder=3, align='edge')
                        # bar.set(bbox=dict(boxstyle='round', color='#d9d9d9'))
                        starttime = str(cls["Start Time"])[0:-4] + ":" + str(cls["Start Time"])[-4:-2]
                        endtime = str(cls["End Time"])[0:-4] + ":" + str(cls["End Time"])[-4:-2]

                        text = f"{cls['Curriculum Item']}  ({starttime} ~ {endtime})\n"
                        text += f"{cls['Full Title']}\n"
                        text += f"{cls['Activity Name']} - Class {cls['Class']}\n"
                        text += f"{cls['Staff Given Name']} {cls['Staff Family Name']}"

                        barheight = bar.patches[0].get_height()
                        barx, bary = bar.patches[0].xy
                        t = plt.text(barx + 0.05, bary + barheight / 2, text, ha='left', va='center', fontsize=7,
                                     wrap=True)
                        wrapsize = (end - start) * 140
                        t._get_wrap_line_width = lambda: wrapsize  # change this counter to adjust wrapping size

                # Name and save
                plt.title(f"{campus} - {dayofweek}\n{self.timeframe}", y=1.07)
                try:
                    mkdir(f"{save_path}/output")
                except FileExistsError as e:
                    pass
                try:
                    mkdir(f'{save_path}/output/{campus}')
                except FileExistsError as e:
                    pass
                plt.savefig(f'{save_path}/output/{campus}/{dayofweek}.png', dpi=200)
                fig.clf()
                plt.close(fig)
            self.tkobj.progbar['value'] += 100 / (len(self.timetable.keys()) - 1)
            self.tkobj.root.update()


if __name__ == '__main__':
    ttv = Application()
