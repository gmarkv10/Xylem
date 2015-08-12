Tree Projection Program written for Town of Amherst:

This program runs with Python and two additional packages:
      'xlrd' and 'xlwt' which can be found at www.python-excel.org

The program is run by double clicking the TreeDataProject Python file

A prompt will ask for a year to project, the prompt only accepts numeric values (2, 5, 10, etc.) and press enter to run the program for that number of years.

The program takes an existing tree data spreadsheet to project.  
IMPORTANT:
     -The spreadsheet must be in the same directory as the program
     -The spreadsheet must be named 'TreeData'
     -The spreadsheet must be a .xls file
     -The spreadsheet must be formatted for iTree

The directory you run the Python file from should also contain a .txt file with the name 'tree_rates', this is where the program gets the rate for each individual tree.

When running, the program will print its progress to the prompt after every thousand entries.
**This may take a while depending on the length of the list and the speed of the computer

When done a new spreadsheet will appear in the directory with the number of years ahead it is projected included in its name.

Press enter to close the prompt. 