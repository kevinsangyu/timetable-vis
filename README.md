# TimeTable Visualiser (ttv)
##### Kevin Yu version 1

### Install required modules:


1. pandas
2. matplotlib
3. openpyxl
4. tkinter

To ensure these modules are installed, run the _setup.py_ file.


### DESCRIPTION AND CONSTRAINTS

This python tool renders timetable images from a single .xlsx file. The Excel spreadsheet file must be in the following 
format:

1. Requires python3
2. Every row of the spreadsheet must be a class/event
3. The spreadsheet must start at A1, meaning that the first row **must** be the column headings
4. The Following headings must be present (Caps Sensitive!):
   1. Curriculum Item (aka Subject Code)
   2. Full Title (aka Subject Name)
   3. Activity Name (aka class type, eg Lecture)
   4. Start Date
   5. Start Time
   6. End Time
   7. Building Id
   8. Room Id
   9. Staff Family Name
   10. Staff Given Name
5. The third last character of the Room Id **must** denote the level of the room, eg OC203 --> Level 2

The resulting images will be outputted into separated folders, under the output folder.

### INSTRUCTIONS
1. Download the files _setup.py_ and _ttv.py_. You only need these 2 files.
2. Run the python file _setup.py_ to ensure all required modules are installed.
3. Run the python file _ttv.py_.
4. This will bring up a console and an open file dialogue window. Navigate to and select your timetable spreadsheet.
5. The program will begin creating the timetable images, storing them in their respective folders in the same directory as the ttv.py file.
6. After it is complete, the console will close on its own, unless there is an error.
