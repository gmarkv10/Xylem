#Xylem
=====

>>Xylem Open Source Projection Software

This is open source software being worked on by Gabe Markarian for Independent Study work being done at UMass Amherst in association with Prof. Rick Harper

In this archetype edition, a useful skeleton is provided in the hopes of making future versions of Xylem easier to understand and replicate.  Most methods and variable are left blank, this version stands as simply a way to grasp one possible working control flow but additions and modifications are encouraged.  

##ASSUMPTIONS ABOUT DEVELOPING FROM ARCHETYPE

  -  Python 2.7 installed along with [xlrd][1] and [xlwt][2] packages for Excel scripting
  -  Excel version post 2003
  -  Simple knowledge of file I/O 
  -  A tree inventory *in .xls format - this means for 2003-2007 Excel versions but later ones can be converted*
  -  The inventory must have fields for a common name and DBH *see lines 148-149 in Xylem_ARCHETYPE.py*

##Xylem_ARCHETYPE.py Documentation

> Line 174 Start of [REPL][3]
  This REPL code is what the user interacts with, if you want to extend anything about the prompts it can be done here.  As you can see there are fucntion calls from code above the REPL, this is where the user interacts with them
  
> Line 14 the `project` function
  This has the loop which iterates down each row of the present tree inventory spreadsheet, each loop calls `project_row` (below).  Line 18 saves the new Excel sheet with the projected data on it, the new file will be called 'Projected' unless otherwise specified.
  
> Line 22 `project_row`
  2 important things happen in projecting each row.  First, the row is marked as valid (see line 106) and then if it is valid, Xylem can decide to do add to the DBH, thus making a projection or simply reprint the row.  The programmer has the choice to make the program not print an invalid row by editing the `print_row` function.
  
> Line 106 `isValidEntry(row)`
  No line in a spreadsheet is guarenteed to have all the proper information for a projection.  This function reads all the fields (e.g. common name and DBH) from the working inventory and assigns their values to in-program variables.  This function is supplemented every time the programmer accounts for a different field (e.g. Condition) (see below for how to add a field).  Conditionals for each field determine if the row can be worked with or not.
  
> Lines 126 & 131: concept of class
   How the programmer wants to handle all the different DBH's they might encounter is here.  Is size on a scale from 1 to 10 or just in inches?  The growth class probably requires more reasoning.  The RATES variable on line 160 is a table that can be filled in with data on how much a tree should grow and a value for growth could be extracted here.  Other methods are equally valid.  
   
###Initializing resouces

Open the config.txt file in /src folder in the archetype and provide these fields
-or-
Run the program and run `setup`

> Lines 142-162
  -  `inventory` is the string in which the inventory name is stored that you are reading from
  -  `common` is the variable that the common name will be stored in as each row is processed
  -  `dbh`    is the variable the dbh will be stored in as each row is processed
  -  `COMMON_COL`&`DBH_COL` are the numbers of the columns in excel the program will look to find those values (see line 90 `setFields`)
  
> Lines 163 - 171
  Open an Excel table to read and write to
  
> Line 43 `init_table`
   Use the rates table to your advantage in memory.  Index rates in way that can be accessed based on other factors about the tree and set up that logic here.
   
   Here is a sample of an init_table already filled in:
   `global INIT_VARS
   
    #keeps init_table from being called before init_vars since it is dependant on the variables it initializes

    if(not INIT_VARS):
    
        print "Startup sequence run incorrectly, please run 'startup' from terminal"

        return

    global RATES

    excel_table = open_workbook(path + 'src\\AnnualPercentageGrowth.xls')

    s1 = excel_table.sheet_by_index(0)
    
    for y in range(100):

        for x in range(13):
        
            RATES[x][y+1] = s1.cell(y, x).value`

   
>Line 49 `init_vars`
   A very important method, this uses the config.txt file in the src directory to read in what columns correspond to what data at program start.  Must be modified for each additional field the progammer considers in their projection.
   
##How to add a new field into your projection program

> If you want an accurate prediction, you'll probably have more parameters than name and dbh, this guide walks you through adding a new paremter.

  -  148 and 155: declare your variables, let's say we're adding a condition varible that's stored as number from 1-100 and it's store in column A in our inventory.  line 150 would be `COND_COL=0` and line 157 would be `cond=0`
  -  Line 51 in `init_vars` would need to include our COND_COL variable in the global declaration.  See [link][4] for why the global keyword is needed.
  -  The config.txt file in ./src must now also have a new line with 0 in it to identify we are looking at the A column.  Then `init_vars` needs a line that looks like `COND_COL = int(obj.readline())` to read it into the program.
  -  `set_fields` on line 90 accesses the cond variable so we'll add it like th dbh variable into the logic there using the helper function `get_cell`
  -  If the field is required for a projection, be sure to make a case for it in `isValidEntry` that way if the field is unusable, the projection won't work.
  -  Finally, prompt the user to set the column in the event they change inventories in `set_vars` on line 68.  After that you are done!
  
  



[1]:https://pypi.python.org/pypi/xlrd
[2]:https://pypi.python.org/pypi/xlwt
[3]:http://en.wikipedia.org/wiki/Read%E2%80%93eval%E2%80%93print_loop
[4]:http://stackoverflow.com/questions/423379/using-global-variables-in-a-function-other-than-the-one-that-created-them
