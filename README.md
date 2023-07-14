# timetable-vis
Kevin Yu v1

Install required modules:
1. pandas
2. matplotlib
3. openpyxl

This python tool renders timetable images from a single .xlsx file. The Excel spreadsheet file must be in the following 
format:

1. Every row must be a class/event
2. The spreadsheet must start at A1, meaning that the first row **must** be the column headings
3. The Following headings must be present (Caps Sensitive!):
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
4. The third last character of the Room Id **must** denote the level of the room, eg OC203 --> Level 2

The resulting images will be outputted into separated folders, under the output folder.