import pandas as pd
from calendar import day_name
import matplotlib.pyplot as plt
from math import modf
# import openpyxl


# Time Table Visualizer
# Kevin Yu


class TimeTableVis(object):
    def __init__(self, excel_path, show_empty_rooms):
        self.excel_path = excel_path
        self.timetable = {}
        self.show_empty_rooms = show_empty_rooms

        self.comprehend_excel_file()

    def comprehend_excel_file(self):
        try:
            file = pd.read_excel(self.excel_path)
        except ValueError as e:
            print("ERROR: Remove any active filters on the excel sheet and retry. \nPress enter to exit.")
            input()
            return
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
        if self.show_empty_rooms:
            self.standardize_rooms()

    def standardize_rooms(self):
        for campus in self.timetable:
            rooms = []
            for dayofweek in self.timetable[campus]:
                for room in self.timetable[campus][dayofweek]:
                    rooms.append(room)
            for dayofweek in self.timetable[campus]:
                for room in rooms:
                    if room not in self.timetable[campus][dayofweek].keys():
                        self.timetable[campus][dayofweek][room] = []

    def text_display_timetable(self):
        for campus in self.timetable:
            print(campus)
            for dayofweek in self.timetable[campus]:
                print("--" + dayofweek)
                for room in self.timetable[campus][dayofweek]:
                    print("----" + room)
                    for cls in self.timetable[campus][dayofweek][room]:
                        print("------", cls["Start Time"])

    def generate_images(self):
        for campus in self.timetable:
            for dayofweek in self.timetable[campus]:
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
                for roomindex in range(len(rooms)):
                    for cls in self.timetable[campus][dayofweek][rooms[roomindex]]:
                        print(f"Plotting campus {campus} on {dayofweek} for {rooms[roomindex]}")
                        start = float(cls["Start Time"]) / 10000
                        end = float(cls["End Time"]) / 10000

                        bar = plt.bar(x=start, height=0.8, width=end-start-0.05, bottom=roomindex+0.6, color='#d9d9d9', ec='black', zorder=3, align='edge')

                        starttime = str(cls["Start Time"])[0:-4] + ":" + str(cls["Start Time"])[-4:-2]
                        endtime = str(cls["End Time"])[0:-4] + ":" + str(cls["End Time"])[-4:-2]

                        text = f"{cls['Curriculum Item']}  ({starttime} ~ {endtime})\n"
                        text += f"{cls['Full Title']}\n"
                        text += f"{cls['Activity Name']} - Class {cls['Class']}\n"
                        text += f"{cls['Staff Given Name']} {cls['Staff Family Name'].upper()}"

                        barheight = bar.patches[0].get_height()
                        barwidth = bar.patches[0].get_width()
                        barx, bary = bar.patches[0].xy
                        t = plt.text(barx + 0.05, bary+barheight/2, text, ha='left', va='center', fontsize=7, wrap=True)
                        t.set(bbox=dict(boxstyle='round'))

                plt.title(campus + " - " + dayofweek, y=1.07)
                plt.savefig('{0}.png'.format(dayofweek), dpi=200)
                fig.clf()
            return


if __name__ == '__main__':
    ttv = TimeTableVis(excel_path="KOI_Class_Report_8496.xlsx", show_empty_rooms=False)
    ttv.generate_images()
