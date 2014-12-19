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
    try:
        print_row(0, False)
        for r in range(1, s.nrows):
            project_row(r, years)
        timestamp = time.strftime('%Y%m%d%H%M%S')
        book.save(path + str(years) + 'Projected'+timestamp+'.xls')
    except AttributeError:
        print "!ERROR: You are using an invalid inventory name"
        print "solution: run `setup` from terminal, being sure to give the right filename"
        print "after completing a setup, be sure to run a `reset` command


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

    
    
INIT_VARS = False
def init_vars():
    global INIT_VARS
    global inventory,SPECIES_COL, COMMON_COL, DBH_COL, HEIGHT_COL, SPREAD_COL, COND_COL, LOC_COL
    global wb, s
    try:
        obj = open(path + 'src\\config.txt', "r")
        inventory = obj.readline()
        inventory = inventory.rstrip('\n')
        wb = open_workbook(path + inventory)
        s = wb.sheet_by_index(0)
        SPECIES_COL = int(obj.readline())
        COMMON_COL  = int(obj.readline())
        DBH_COL     = int(obj.readline())
        HEIGHT_COL  = int(obj.readline())
        SPREAD_COL  = int(obj.readline())
        COND_COL    = int(obj.readline())
        LOC_COL     = int(obj.readline())
        obj.close()
        INIT_VARS = True
    except IOError:
        print "!ERROR: " + inventory + " does not exist in this folder"
        print "solution: run `setup` from terminal, being sure to give the right filename"
        print "after completing a setup, be sure to run a `reset` command
    
def set_vars():
    global path
    print "doing a lot of important settings work"
    obj = open(path + 'src\\config.txt', "w")
    ipath = raw_input("What is the inventory name? ")
    obj.write(ipath + '.xls\n')
    print "Columns can be given as upper or lower case"
    setCol = raw_input("What column is the SPECIES NAME? ")
    obj.write(ctoi(setCol)+'\n')
    setCol = raw_input("What column is the COMMON NAME? ")
    obj.write(ctoi(setCol)+'\n')
    setCol = raw_input("What column is the DBH? ")
    obj.write(ctoi(setCol)+'\n')
    setCol = raw_input("What column is the HEIGHT? ")
    obj.write(ctoi(setCol)+'\n')
    setCol = raw_input("What column is the SPREAD? ")
    obj.write(ctoi(setCol)+'\n')
    setCol = raw_input("What column is the CONDITION? ")
    obj.write(ctoi(setCol)+'\n')
    setCol = raw_input("What column is the LOCATION RATING? ")
    obj.write(ctoi(setCol)+'\n')
    #init_vars()

def ctoi(c):
    c = c.lower()
    return str(ord(c) - 97)


def init_table():
    global INIT_VARS
    #keeps init_table from being called before init_vars since it is dependant on the variables it initializes
    if(not INIT_VARS):
        print "Startup sequence run incorrectly, please run `startup` from terminal"
        return
    global RATES
    excel_table = open_workbook(path + 'src\\AnnualPercentageGrowth.xls')
    s1 = excel_table.sheet_by_index(0)
    for y in range(100):
        for x in range(13):
            RATES[x][y+1] = s1.cell(y, x).value
        
    #print s1.cell(99,12).value

    
       

def getCell(r, c):
    return s.cell(r, c).value

def setFields(row):
    global loc, species, common, dbh, height, spread, cond
    try:
        species = getCell(row, SPECIES_COL)
        common  = getCell(row, COMMON_COL)
        dbh     = float(getCell(row, DBH_COL))
        height  = float(getCell(row, HEIGHT_COL))
        spread  = float(getCell(row, SPREAD_COL))
        cond    = float(getCell(row, COND_COL))
        loc     = float(getCell(row, LOC_COL))
    except ValueError:
        #On excpetion: let isValid entry decide validity of raw
        #cell value, trivially it will return false
        species = getCell(row, SPECIES_COL)
        common  = getCell(row, COMMON_COL)
        dbh     = getCell(row, DBH_COL)
        height  = getCell(row, HEIGHT_COL)
        spread  = getCell(row, SPREAD_COL)
        cond    = getCell(row, COND_COL)
        loc     = getCell(row, LOC_COL)
    

def test():
    i = 0
    while(i < 5):
        s = raw_input(">>> ")
        print ctoi(s)
        i += 1
    #init_vars()
    #project_row(4, 5)
    #book.save(path + 'Projecte.xls')
    
    
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

def reset():
    init_vars()
    init_table()

#########################    
#Pre-REPL instantiations#
#########################
print "WELCOME TO XYLEM Version 0.0.2 for the UMASS CAMPUS"
print "getting things ready..."
#Startup sequence

inventory = ""


SPECIES_COL = -1
COMMON_COL  = -2
DBH_COL     = -3
HEIGHT_COL  = -4
SPREAD_COL  = -5
COND_COL    = -7
LOC_COL     = -9

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
    elif(cmd == 'reset'):
        reset()
    elif(cmd == 'setup'):
        set_vars()
    else:
        print "USAGE: 'grow' -- grow the inventory by some number of years"
        print "       'quit' or 'q' -- exit the program"
        print "       'setup' -- prompt to reset path and columns for inventory"
        print "       'reset' -- restablish connection with growth table and inventory, use after setup"
        
