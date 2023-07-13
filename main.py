import pandas as pd
import datetime
import calendar

if __name__ == '__main__':
    book = pd.read_excel("KOI_Class_Report_8496.xlsx")
    timetable = {}
    for row in book.iloc():
        date = row["Start Date"]
        dayoftheweek = calendar.day_name[date.weekday()]
        if row["Building Id"] not in timetable.keys():
            timetable[row["Building Id"]] = {'Monday': {}, 'Tuesday': {}, 'Wednesday': {},
                                             'Thursday': {}, 'Friday': {}, 'Saturday':{},
                                             'Sunday': {}}
            timetable[row["Building Id"]][dayoftheweek][row["Room Id"]] = [row]
        elif row["Room Id"] not in timetable[row["Building Id"]][dayoftheweek].keys():
            timetable[row["Building Id"]][dayoftheweek][row["Room Id"]] = [row]
        else:
            timetable[row["Building Id"]][dayoftheweek][row["Room Id"]].append(row)
    for campus in timetable:
        print(campus)
        for dayofweek in timetable[campus]:
            print("--" + dayofweek)
            for room in timetable[campus][dayofweek]:
                print("----" + room)
                for cls in timetable[campus][dayofweek][room]:
                    print("------", cls["Start Time"])
