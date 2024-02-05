# TimeTable Visualiser (ttv)
##### Kevin Yu version 2

### Install required modules:


1. pandas
2. matplotlib
3. openpyxl
4. tkinter

To ensure these modules are installed, run the _setup.py_ file.


### DESCRIPTION AND CONSTRAINTS

This python tool renders timetable images from a single .xlsx file. The following constrains apply:

1. Requires python3
2. Every row of the spreadsheet must be a class/event
3. The spreadsheet must start at A1, meaning that the first row **must** be the column headings
4. The Following headings must be present in the spreadsheet (Caps Sensitive!):
   1. Curriculum Item (aka Subject Code)
   2. Full Title (aka Subject Name)
   3. Activity Name (aka class type, e.g. Lecture)
   4. Start Date (Used to calculate the day of week of class)
   5. Start Time (Class start time)
   6. End Time (Class end time)
   7. Building Id
   8. Room Id
   9. Staff Family Name
   10. Staff Given Name
   11. Year
   12. Study Period - Code (Term/Semester number, Used for timetable title)
   13. Study Period - Date (Date period of Term/Semester, Used for timetable title)
5. The third last character of the Room Id **must** denote the level of the room, e.g. OC203 --> Level 2

The resulting images will be outputted into separated folders, under the output folder. The output folder will be created in the same directory as the selected spreadsheet.

### INSTRUCTIONS
1. Download the files _setup.py_ and _ttv.pyw_. You only need these 2 files.
2. Run the python file _setup.py_ to ensure all required modules are installed.
3. Run the python file _ttv.pyw_.
4. Follow the prompts to generate your timetable.

Note that if you want to run this code on Linux or macOS, it should be functional if you change the file extension from .pyw to .py.
