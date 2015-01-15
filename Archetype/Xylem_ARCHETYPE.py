#Xylem open source tree projection for UMASS INVENTORY 3.14.14
from xlrd import open_workbook, empty_cell
from xlwt import Workbook
import os
import time


path = os.getcwd() + '\\'
#######################
#Function delcarations#
#######################

#contains the loop that goes over all inventory entries, stamps each newly projected file with years ahead and time created.
def project(years):
    print_row(0, False)
    for r in range(1, s.nrows):
        project_row(r, years)
    book.save(path + 'Projected.xls')


#uses helper functions to project an entry, and print row to write to the projected file
def project_row(row, years):
    global dbh
    if(isValidEntry(row)):
        print_row(row, True)
        
    else:
        print_row(row, False)

#print_row simply reprints a row to the new file if not all the requirements for projection are met
#otherwise it prints the row with the projected dbh made from project_row
def print_row(row, valid):
    for c in range(s.ncols):
        if(c == DBH_COL):
            if(valid):
                sheet1.write(row, DBH_COL, dbh)
            else:
                sheet1.write(row, DBH_COL, getCell(row, DBH_COL))
        else:
            sheet1.write(row, c, s.cell(row, c).value)

#Reads in the excel version of Kim Coder's article into an array which is indexed in O(1) time, increasing efficiency
def init_table():
    global RATES
    
    

#read in the info from config.txt, one note, this needs to be run before init_table()
def init_vars():
   
    global COMMON_COL, DBH_COL
    global wb, s
    try:
        obj = open(path + 'src\\config.txt', "r")
        inventory = obj.readline()
        inventory = inventory.rstrip('\n')
        wb = open_workbook(path + inventory)
        s = wb.sheet_by_index(0)
        COMMON_COL  = int(obj.readline())
        DBH_COL     = int(obj.readline())
        obj.close()
    except IOError:
        print "!ERROR: " + inventory + " does not exist in this folder"
        print "solution: run `setup` from terminal, being sure to give the right filename"
        print "after completing a setup, be sure to run a `reset` command"

#overwrite the program info in config.txt for a new inventory
def set_vars():
    global path
    print "doing a lot of important settings work"
    obj = open(path + 'src\\config.txt', "w")
    ipath = raw_input("What is the inventory name? ")
    obj.write(ipath + '.xls\n')
    print "Columns can be given as upper or lower case"
    setCol = raw_input("What column is the COMMON NAME? ")
    obj.write(ctoi(setCol)+'\n')
    setCol = raw_input("What column is the DBH? ")
    obj.write(ctoi(setCol)+'\n')
    #init_vars()





    
       

#called in isValidEntry() since that is called before each projection,
#thereby resetting the fields
def setFields(row):
    global loc, species, common, dbh, height, spread, cond
    try:
        common  = getCell(row, COMMON_COL)
        dbh     = float(getCell(row, DBH_COL))
    except ValueError:
        #On excpetion: let isValid entry decide validity of raw
        #cell value, trivially it will return false
        common  = getCell(row, COMMON_COL)
        dbh     = getCell(row, DBH_COL)
       


    
    
#tests for validty of the line, also calls setFields!
def isValidEntry(row):
    if(row > s.nrows):
        return False
    setFields(row)
    if(common  == empty_cell.value):
        return False
    if(dbh  == empty_cell.value or dbh < 0 or type(dbh) is not float):
        return False
    return True

#easy access to cell vals
def getCell(r, c):
    return s.cell(r, c).value

#used in set_vars function to attain the numerical location of the column from the letter notation
def ctoi(c):
    c = c.lower()
    return str(ord(c) - 97)


def getDbhClassFromDBH(dbh):
    return dbh
    

#Called by project_row() returning the approriate index for RATES table or otherwise
def getGrowthClassFromCond(params):
    return params


#Test function, called from terminal 
def test():
    print 'In-house tester method, called by `test` from terminal'

#########################    
#Pre-REPL instantiations#
#########################
print "WELCOME TO YOUR OWN VERSION OF XYLEM"
print "getting things ready..."
#Startup sequence

inventory = ""

COMMON_COL  = -2
DBH_COL     = -3


#(above and below) python vars for working with fields  
  

common = ""
dbh    = -1


#2D growth increment array that is initialized with `init_table` above
RATES = [[]]


wb = ''#= open_workbook(path + inventory)
s = ''#= wb.sheet_by_index(0)      These assignments can now be found in init_vars()
##############
init_vars()  #Initializing those fields 
##############
book = Workbook()
sheet1 = book.add_sheet('Sheet 1')

init_table()

print "all set, type any non-command for usage"
#######################################
#REPL code, will run until termination#
#######################################
cmd = ""

while(cmd != 'quit' and cmd != 'q'):
    cmd = raw_input("> ")
    if(cmd == 'test'):
        test()
    elif(cmd == 'quit' or cmd == 'q'):
        print " "
        #I'll show myself out
    elif(cmd == 'grow' or cmd == 'project'):
        print " How many years of growth?"
        years = input(">> ")
        project(years)
    elif(cmd == 'setup'):
        set_vars()
    else:
        print "USAGE: 'grow' -- grow the inventory by some number of years"
        print "       'quit' or 'q' -- exit the program"
        print "       'setup' -- prompt to reset path and columns for inventory"
        
