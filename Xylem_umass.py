#Xylem open source tree projection for UMASS INVENTORY 3.14.14
from xlrd import open_workbook, empty_cell
from xlwt import Workbook
import os
import time

path = os.getcwd() + '\\'
#######################
#Function delcarations#
#######################

def project(years):
    print_row(0, False)
    for r in range(1, s.nrows):
        project_row(r, years)
    timestamp = time.strftime('%Y%m%d%H%M%S')
    book.save(path + str(years) + 'Projected'+timestamp+'.xls')


def project_row(row, years):
    global dbh
    global spread, height, cond, loc
    if(isValidEntry(row)):
        for y in range(years):
            dbhClass = int(getDbhClassFromDBH(dbh))
            gClass   = int(getGrowthClassFromCond(spread, height, cond, loc))
            dbh      = dbh + (dbh*RATES[gClass][dbhClass])
        print_row(row, True)
        
    else:
        print_row(row, False)


def print_row(row, valid):
    for c in range(s.ncols):
        if(c == DBH_COL):
            if(valid):
                sheet1.write(row, DBH_COL, dbh)
            else:
                sheet1.write(row, DBH_COL, getCell(row, DBH_COL))
        else:
            sheet1.write(row, c, s.cell(row, c).value)
    #above code, just prints the line,
    #the below code writes over the old dbh and puts in the new one for a valid enrty

    
    
    
    
def init_table():
    global RATES
    excel_table = open_workbook(path + 'AnnualPercentageGrowth.xls')
    s1 = excel_table.sheet_by_index(0)
    for y in range(100):
        for x in range(13):
            RATES[x][y+1] = s1.cell(y, x).value
        
    #print s1.cell(99,12).value
       

def getCell(r, c):
    return s.cell(r, c).value

def setFields(row):
    global loc, species, common, dbh, height, spread, cond
    species = getCell(row, SPECIES_COL)
    common  = getCell(row, COMMON_COL)
    dbh     = float(getCell(row, DBH_COL))
    height  = float(getCell(row, HEIGHT_COL))
    spread  = float(getCell(row, SPREAD_COL))
    cond    = float(getCell(row, COND_COL))
    loc     = float(getCell(row, LOC_COL))
    
    

def test():
    
    project_row(3, 2)
    project_row(4, 3)
    book.save(path + 'Projecte.xls')
    
    
#tests for validty of the line, also calls setFields!
def isValidEntry(row):
    if(row > s.nrows):
        return False
    setFields(row)
    if(species == empty_cell.value):
        return False
    if(common  == empty_cell.value):
        return False
    #may be changed later to handle all dbh's
    if(dbh  == empty_cell.value or dbh < 6 or type(dbh) is not float):
        return False
    if( height  == empty_cell.value or height == 0 or type(height) is not float):
        return False
    if( spread  == empty_cell.value or spread == 0 or type(spread) is not float):
        return False
    if( cond  == empty_cell.value or type(cond) is not float):
        return False
    if( loc  == empty_cell.value or type(loc) is not float):
        return False
    
    return True

def getDbhClassFromDBH(dbh):
    if(dbh > 0 and dbh <= 40):
        return dbh
    elif( dbh > 40 and dbh <= 100):
        return incofFive(dbh)  #only listed in incs of 5 after dbh>40, see Kim Coder's article
    #an entry greater than 100 ( the highest DBH in Coder's chart ) will get
    #pidgeon holed into 100
    elif( dbh > 100 ):
        return 100
    else:
        return "!CRITICAL: Invalid DBH"

def incofFive(num):
    mod = num % 5
    if(mod == 0):
        return num
    elif( mod >= 3):
        return num + (5 - mod)
    else:
        return num - mod

def getGrowthClassFromCond(spread, height, cond, loc):
    #BY INDEX of column vs growth increments per inch:
    # 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9  | 10 | 11 | 12
    #---|---|---|---|---|---|---|---|---|----|----|----|----
    # 1 |1.5| 2 |2.5| 3 | 4 | 5 |7.5|10 |12.5| 15 |17.5| 20
    index = 0
    if(spread/height >= 4):
        index += 4
    elif(spread/height >= 3):
        index += 2
    elif(spread/height >= 2):
        index += 1
    #MAX possible 4
    if(cond > 75):
        index += 0
    elif(cond > 50 ):
        index += 1
    elif(cond > 25):
        index += 2
    elif(cond > 0):
        index += 3
    else:
        index += 4
    #MAX possible 8
    if(loc > 75):
        index += 0
    elif(loc > 50 ):
        index += 1
    elif(loc > 25):
        index += 2
    elif(loc > 0):
        index += 3
    else:
        index += 4
    #MAX possible 12
    
    return index


#########################    
#Pre-REPL instantiations#
#########################
print "WELCOME TO XYLEM Version 0.0.1 for the UMASS CAMPUS"
print "getting things ready..."
#Startup sequence




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
RATES = [[0 for y in range(101)] for x in range(13)]

wb = open_workbook(path + 'umass.xls')
s = wb.sheet_by_index(0)

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
        project(years)
    else:
        print "USAGE: 'grow' -- grow the inventory by some number of years"
        print "       'quit' or 'q' -- exit the program"
        print "       'setcol <species> <common> <dbh> <height> <spread> <condition> <loc rating> --"
        print "             tell the program which columns the fields are in"
