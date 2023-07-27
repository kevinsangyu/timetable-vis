import tkinter as tk
from calendar import day_name
from os import mkdir
from tkinter import ttk, messagebox, Toplevel
from tkinter.filedialog import askopenfilename
from textwrap import wrap
from subprocess import Popen
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
        self.imagepath_label = ttk.Label(self.root, text="Open logo image\n(Recommended size 239x92)", width=25,
                                         anchor='w')
        self.filepath_button = ttk.Button(self.root, text="Open file", command=self.open_spreadsheet, width=11)
        self.imagepath_button = ttk.Button(self.root, text="Open Image", command=self.open_image, width=11)
        self.generate_button = ttk.Button(self.root, text="Generate!", command=self.generate, width=11,
                                          state='disabled')
        self.exit_button = ttk.Button(self.root, text="Exit", command=self.root.destroy)
        self.child = None  # to silence IDE warnings
        self.progbar = None
        self.prog_label = None
        self.detail_label = None

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
            if self.imagepath_label['text'] != '' and self.imagepath_label['text'] != 'Open logo image':
                self.generate_button["state"] = 'enable'

    def open_image(self):
        filetypes = (('PNG Image', '*.png'), ('All files', '*.*'))
        image_path = askopenfilename(filetypes=filetypes, initialdir=".\\", title="Locate logo image")
        if image_path:
            self.imagepath_label.config(text=image_path)
            if self.filepath_label['text'] != '' and self.filepath_label['text'] != 'Open Spreadsheet file':
                self.generate_button["state"] = 'enable'

    def generate(self):
        excel_path = self.filepath_label.cget('text')
        image_path = self.imagepath_label.cget('text')
        self.generate_button['state'] = 'disable'
        ttv_client = TimeTableVis(self, excel_path, image_path)
        save_path = ttv_client.save_path + "/output"
        save_path = save_path.replace("/", "\\")
        Popen(f'explorer "{save_path}"')

    def error(self, status):
        messagebox.showerror('Error', "Failed Excel spreadsheet comprehension!\n"
                                      "Try removing any active filters in the file, and make sure that the first "
                                      "column heading is at cell A1. Note that there may be hidden columns or rows at "
                                      f"the start of excel files. \n\nError: {status}")
        self.root.update()
        self.child.destroy()
        self.generate_button['state'] = 'enable'

    def prog_window(self):
        self.child = Toplevel(self.root)
        self.child.geometry("300x120")
        self.child.resizable(False, False)
        self.child.title("TTV - Generating Images...")
        self.progbar = ttk.Progressbar(self.child, orient='horizontal', mode='determinate', length=200)
        self.prog_label = ttk.Label(self.child, text="Initialising...", width=45, anchor='w')
        self.detail_label = ttk.Label(self.child, text="", width=45, anchor='w')
        self.progbar.grid(row=0, column=0, columnspan=2, padx=10, pady=20)
        self.prog_label.grid(row=1, column=0, columnspan=2, padx=10)
        self.detail_label.grid(row=2, column=0, columnspan=2, padx=10)
        self.root.update()

    def close(self):
        self.root.destroy()
    # todo make separate methods to update the progress bar instead of accessing this classes' members
    # that would make it more OOP but it is a pain


class TimeTableVis(object):
    def __init__(self, tkobj, excel_path, image_path):
        self.excel_path = excel_path
        self.timetable = {}
        self.timeframe = ""  # year and term number
        self.tkobj = tkobj
        self.save_path = "/".join(excel_path.split("/")[0:-1])

        tkobj.prog_window()
        self.comprehend_excel_file()
        self.generate_images(self.save_path, image_path)
        messagebox.showinfo('Complete', 'Timetable generation has been completed.\nIt is located in the same directory'
                                        ' as the spreadsheet file, under the folder "output". After this window is '
                                        'closed file explorer will open it.')
        self.tkobj.close()

    def comprehend_excel_file(self):
        try:
            file = pd.read_excel(self.excel_path)
        except ValueError:
            self.tkobj.error("Failed at excel sheet comprehension (pandas). Check format of excel file.")
            return
        self.tkobj.progbar['value'] = 0
        self.tkobj.prog_label.config(text="Comprehending excel sheet file...")
        required = ['Year', 'Study Period - Code', 'Start Date', 'Building Id', 'Room Id', 'Start Time', 'End Time']
        # only for functionally required columns, not visually required columns. i.e. info to draw, not label
        for item in required:
            if item not in list(file):
                self.tkobj.error(f"Failed at excel sheet comprehension, required column '{item}' not present. Check "
                                 f"for spelling errors.")
                return
        self.timeframe = f"{file['Year'][0]} - {file['Study Period - Code'][0]}"
        for row in file.iloc():
            skip = False
            for item in required:
                if row[item] != row[item]:  # check if NaN (float representation of Not-a-Number)
                    self.tkobj.prog_label.config(text="Skipping row with no value in functionally required columns")
                    skip = True
            if skip:
                continue
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
        make sure every classroom in the campus is present in the timetable. Otherwise, the timetable would only display
        classrooms which have classes/events on that day.

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
            # standardize rooms
            rooms = []
            for dayofweek in self.timetable[campus]:
                for room in self.timetable[campus][dayofweek]:
                    rooms.append(room)
            for dayofweek in self.timetable[campus]:
                for room in rooms:
                    if room not in self.timetable[campus][dayofweek].keys():
                        self.timetable[campus][dayofweek][room] = []
            # standardize levels
            levels = set()
            for room in rooms:
                levels.add(str(room)[-3])
                # third last character of the room id denotes the level of the room. (See method description)
            if len(levels) > 1:
                self.tkobj.detail_label.config(text=f"Multiple levels found on campus {campus}")
                self.tkobj.root.update()
                for level in levels:
                    self.tkobj.detail_label.config(text=f"Adding classes for {campus} Level {level}")
                    self.tkobj.root.update()
                    self.timetable[f'{campus} Level {level}'] = {'Monday': {}, 'Tuesday': {}, 'Wednesday': {},
                                                                 'Thursday': {}, 'Friday': {}, 'Saturday': {},
                                                                 'Sunday': {}}
                    for room in rooms:
                        if str(room)[-3] == level:
                            for dayofweek in self.timetable[f'{campus} Level {level}']:
                                self.timetable[f'{campus} Level {level}'][dayofweek][room] = \
                                    [cls for cls in self.timetable[campus][dayofweek][room] if
                                     cls['Room Id'][-3] == level]

                del self.timetable[campus]
            self.tkobj.progbar['value'] += 100 / (len(copytable.keys()) - 1)
            self.tkobj.root.update()

    def generate_images(self, save_path, image_path):
        """
        Generates the timetable images.
        The times are hard coded between 9:00 and 21:00. The room numbers are added automatically depending on the rooms
        in each campus.
        The label will display the subject code, subject name, activity name and number, staff family and given name.
        The logo is inserted on the bottom left of the image.
        The title is CAMPUS Level - Day of week, Year Term number.
        """
        self.tkobj.prog_label.config(text="Generating timetables...")
        self.tkobj.detail_label.config(text="")
        self.tkobj.progbar['value'] = 0
        self.tkobj.root.update()
        for campus in self.timetable:
            for dayofweek in self.timetable[campus]:
                self.tkobj.detail_label.config(text=f"Plotting campus {campus} for {dayofweek}")
                fig, ax = plt.subplots(figsize=(12, 7))
                rooms = []
                for room in self.timetable[campus][dayofweek]:
                    rooms.append(room)
                rooms.sort(reverse=True)

                # Axis formatting
                ax.xaxis.grid(zorder=0)  # z order is to hide it behind bars/classes
                ax.set_ylim(0.3, len(rooms) + 0.7)
                ax.set_xlim(8.9, 21.9)
                ax.set_yticks(range(1, len(rooms) + 1))
                ax.set_yticklabels(rooms)
                ax.set_xticks(range(9, 23))
                ax.set_xticklabels(["09:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00",
                                    "18:00", "19:00", "20:00", "21:00", "22:00"])
                ax.set_xlabel('Time')

                # Adding logo
                # The size of the image is designed to be around (239x92), anything bigger or smaller will display
                # as bigger or smaller in the output, which could either be unreadable or hide parts of the timetable.
                logo = image.imread(image_path)
                fig.figimage(logo, 10, 10)

                # Plotting and labelling classes/events
                for roomindex in range(len(rooms)):
                    if roomindex != 0:
                        # line in between boxes, aka y grid
                        ax.axline((1, roomindex + 0.5), (2, roomindex + 0.5), color='grey', linewidth=0.5)
                    for cls in self.timetable[campus][dayofweek][rooms[roomindex]]:
                        start = float(cls["Start Time"]) / 10000
                        end = float(cls["End Time"]) / 10000

                        # Plot the class/event
                        bar = plt.bar(x=start, height=0.8, width=end - start - 0.05, bottom=roomindex + 0.6,
                                      color='#d9d9d9', ec='black', zorder=3, align='edge')
                        # bar.set(bbox=dict(boxstyle='round', color='#d9d9d9'))

                        # Label the class/event
                        starttime = str(int(cls["Start Time"]))[0:-4] + ":" + str(int(cls["Start Time"]))[-4:-2]
                        endtime = str(int(cls["End Time"]))[0:-4] + ":" + str(int(cls["End Time"]))[-4:-2]
                        text = f"{cls['Curriculum Item']}  ({starttime} ~ {endtime})\n"
                        text += "\n".join(wrap(cls['Full Title'], width=12 * int(end - start))) + "\n"
                        text += f"{cls['Activity Name']} - {cls['Class']}\n"
                        text += f"{cls['Staff Given Name']} {cls['Staff Family Name']}"

                        barheight = bar.patches[0].get_height()
                        barx, bary = bar.patches[0].xy
                        plt.text(barx + 0.05, bary + barheight / 2, text, ha='left', va='center', fontsize=7)

                # Name and save
                plt.title(f"{campus} - {dayofweek}\n{self.timeframe}", y=1.07)
                try:
                    mkdir(f"{save_path}/output")
                except FileExistsError:
                    pass
                try:
                    mkdir(f'{save_path}/output/{campus}')
                except FileExistsError:
                    pass
                plt.savefig(f'{save_path}/output/{campus}/{dayofweek}.png', dpi=200)
                # dpi relates to the size of the image, which normally wouldn't matter, but since an image of variable
                # size is being inserted, with the sample image it's designed for (239x92 pixels), 200 is a good size.
                fig.clf()
                plt.close(fig)
                self.tkobj.progbar['value'] += 100 / len(self.timetable.keys()) / 7
                self.tkobj.root.update()


if __name__ == '__main__':
    ttv = Application()
