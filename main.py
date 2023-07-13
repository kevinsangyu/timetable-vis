import pandas as pd
import calendar
import matplotlib.pyplot as plt

# Time Table Visualizer
# Kevin Yu


class TimeTableVis(object):
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.timetable = {}
        self.comprehend_excel_file()

    def comprehend_excel_file(self):
        file = pd.read_excel(self.excel_path)
        for row in file.iloc():
            date = row["Start Date"]
            dayoftheweek = calendar.day_name[date.weekday()]
            if row["Building Id"] not in self.timetable.keys():
                self.timetable[row["Building Id"]] = {'Monday': {}, 'Tuesday': {}, 'Wednesday': {},
                                                 'Thursday': {}, 'Friday': {}, 'Saturday': {},
                                                 'Sunday': {}}
                self.timetable[row["Building Id"]][dayoftheweek][row["Room Id"]] = [row]
            elif row["Room Id"] not in self.timetable[row["Building Id"]][dayoftheweek].keys():
                self.timetable[row["Building Id"]][dayoftheweek][row["Room Id"]] = [row]
            else:
                self.timetable[row["Building Id"]][dayoftheweek][row["Room Id"]].append(row)

    def text_display_timetable(self):
        for campus in self.timetable:
            print(campus)
            for dayofweek in self.timetable[campus]:
                print("--" + dayofweek)
                for room in self.timetable[campus][dayofweek]:
                    print("----" + room)
                    for cls in self.timetable[campus][dayofweek][room]:
                        print("------", cls["Start Time"])


if __name__ == '__main__':
    ttv = TimeTableVis("KOI_Class_Report_8496.xlsx")
    ttv.text_display_timetable()
