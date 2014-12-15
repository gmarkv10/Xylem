#Xylem open source tree projection for UMASS INVENTORY 3.14.14
from xlrd import open_workbook, empty_cell
from xlwt import Workbook
import os

path = os.getcwd() + '\\'
#######################
#Function delcarations#
#######################


#def grow_out(years):
    
    
def init_table():
    global RATES
    excel_table = open_workbook(path + 'AnnualPercentageGrowth.xls')
    s1 = excel_table.sheet_by_index(0)
    for iclass in range(47):
        for dbh in range(13):
            RATES[iclass][dbh] = s1.cell(1+ iclass,1 + dbh).value
       

def getCell(r, c):
    return s.cell(r, c).value

def setFields(row):
    global loc
    global species
    global common
    global dbh
    global height
    global spread
    global cond
    species = getCell(row, SPECIES_COL)
    common  = getCell(row, COMMON_COL)
    dbh     = getCell(row, DBH_COL)
    height  = getCell(row, HEIGHT_COL)
    spread  = getCell(row, SPREAD_COL)
    cond    = getCell(row, COND_COL)
    loc     = getCell(row, LOC_COL)
    
    

def test():
    global RATES
    init_table()
    print RATES

#tests for validty of the line, also calls setFields!
def isValidEntry(row):
    if(row > s.nrows):
        return False
    setFields(row)
    if(species == empty_cell.value):
        return False
    if(common  == empty_cell.value):
        return False
    if(dbh  == empty_cell.value or dbh == 0):
        return False
    if( height  == empty_cell.value or height == 0):
        return False
    if( spread  == empty_cell.value or spread == 0):
        return False
    if( cond  == empty_cell.value):
        return False
    if( loc  == empty_cell.value):
        return False
    return True
#########################    
#Pre-REPL instantiations#
#########################

#Version 0.0.2 will read these in from an external file for adjustabililty

SPECIES_COL = 1
COMMON_COL  = 2
DBH_COL     = 3
HEIGHT_COL  = 4
SPREAD_COL  = 5
COND_COL    = 7
LOC_COL     = 9

#(above and below) python vars for working with fields  
  
species = ""
common = ""
dbh    = -1
height = -1
spread = -1
cond   = -1
loc    = -1

#2D growth increment array that is initialize with `init_table` above
#accessed by RATES[<incremenent class>][<dbh>]
#Read in from Kim D. Coder's paper on Annual Percentage Growth in xls form
RATES = [[0 for i in range(13)] for i in range(47)]

wb = open_workbook(path + 'umass.xls')
s = wb.sheet_by_index(0)

book = Workbook()
sheet1 = book.add_sheet('Sheet 1')



#######################################
#REPL code, will run until termination#
#######################################
cmd = ""
print "WELCOME TO XYLEM Version 0.0.1 for the UMASS CAMPUS"
while(cmd != 'quit' and cmd != 'q'):
    cmd = raw_input("> ")
    if(cmd == 'proj'):
        print "Doing a lot of important projetion work"
    elif(cmd == 'test'):
        test()
    elif(cmd == 'quit' or cmd == 'q'):
        print " "
        #I'll show myself out
    elif(cmd == 'grow' or cmd == 'project'):
        print " How many years of growth?"
        years = input(">> ")
        print years
        #grow_out(years)
    else:
        print "USAGE: 'grow' -- grow the inventory by some number of years"
        print "       'quit' or 'q' -- exit the program"
        print "       'setcol <species> <common> <dbh> <height> <spread> <condition> <loc rating> --"
        print "             tell the program which columns the fields are in"