import pandas as pd
from calendar import day_name
import matplotlib.pyplot as plt
from os import mkdir


# Time Table Visualizer
# Kevin Yu


class TimeTableVis(object):
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.timetable = {}

        self.comprehend_excel_file()

    def comprehend_excel_file(self):
        try:
            file = pd.read_excel(self.excel_path)
        except ValueError as e:
            print("ERROR: Remove any active filters on the excel sheet and retry. \nPress enter to exit.")
            input()
            return
        print("Comprehending excel sheet file")
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
                print(f"Multiple levels found on campus {campus}")
                for level in levels:
                    print(f"Adding classes for f'{campus}L{level}'")
                    self.timetable[f'{campus}L{level}'] = {'Monday': {}, 'Tuesday': {}, 'Wednesday': {},
                                                           'Thursday': {}, 'Friday': {}, 'Saturday': {},
                                                           'Sunday': {}}
                    for room in rooms:
                        if str(room)[-3] == level:
                            for dayofweek in self.timetable[f'{campus}L{level}']:
                                self.timetable[f'{campus}L{level}'][dayofweek][room] = [cls for cls in
                                                                                        self.timetable[campus][
                                                                                            dayofweek][room] if
                                                                                        cls['Room Id'][-3] == level]

                del self.timetable[campus]

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

    def generate_images(self):
        """
        Actually does the generation of timetable images
        """
        for campus in self.timetable:
            for dayofweek in self.timetable[campus]:
                print(f"Plotting campus {campus} for {dayofweek}")
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
                ax.set_xlabel('Time')

                # Plotting Classes
                for roomindex in range(len(rooms)):
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
                        wrapsize = (end - start) * 400
                        t._get_wrap_line_width = lambda: wrapsize  # change this counter to adjust wrapping size

                # Name and save
                plt.title(campus + " - " + dayofweek, y=1.07)
                try:
                    mkdir("output")
                except FileExistsError as e:
                    pass
                try:
                    mkdir(f'output/{campus}')
                except FileExistsError as e:
                    pass
                plt.savefig(f'output/{campus}/{dayofweek}.png', dpi=600)
                fig.clf()
                plt.close(fig)


if __name__ == '__main__':
    ttv = TimeTableVis(excel_path="KOI_Class_Report_8496.xlsx")
    ttv.generate_images()
