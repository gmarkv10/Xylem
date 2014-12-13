#Xylem open source tree projection for UMASS INVENTORY 3.14.14
from xlrd import open_workbook, empty_cell
from xlwt import Workbook
import os

path = os.getcwd() + '\\'
#######################
#Function delcarations#
#######################
def myFunc():
    for item in range(s.ncols):
        sheet1.write(1, item, s.cell(1, item).value)

    book.save(path + 'projected.xls')

    print "we good"

#def grow_out(years):
    
#def init_table():

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
    print isValidEntry(656456)

#def isValidEntry(species, common, dbh, height, spread, cond, loc):
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
        myFunc()
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
        print 'USAGE: proj'
