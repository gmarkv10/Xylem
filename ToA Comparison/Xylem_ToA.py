#Xylem open source tree projection for UMASS INVENTORY 3.14.14
from xlrd import open_workbook, empty_cell
from xlwt import Workbook
import os
import time
import random


path = os.getcwd() + '\\'
#######################
#Function delcarations#
#######################
prinfo = True
loop = 0
#contains the loop that goes over all inventory entries, stamps each newly projected file with years ahead and time created.
def project(years):
    global prinfo
    try:
        print_row(0, False)
        for r in range(1, s.nrows):
            if(r > 5):
                prinfo = False
            else:
                print "---------"
            project_row(r, years)
        timestamp = time.strftime('%Y%m%d%H%M%S')
        book.save(path + str(years) + 'Projected'+timestamp+'.xls')
        print inventory + " projected successfully for " + str(years) + " years"
    except AttributeError:
        print "!ERROR: You are using an invalid inventory name"
        print "solution: run `setup` from terminal, being sure to give the right filename"
        print "after completing a setup, be sure to run a `reset` command"

#uses helper functions to project an entry, and print row to write to the projected file
def project_row(row, years):
    global dbh, loop
    global spread, height, cond, loc, space
    if(isValidEntry(row)):
        for y in range(years):
            loop = y
            dbhClass = int(getDbhClassFromDBH(dbh))
            gClass   = int(getGrowthClassFromCond(spread, height, cond, loc, space))
            #print (RATES[gClass][dbhClass])
            #print "Row " + str(row) + " -- Class " + str(gClass)
            dbh      = dbh + (dbh*RATES[gClass][dbhClass])
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
    
INIT_VARS = False   #Failsafe to make sure the variables are initialized at run time
#read in the info from config.txt, one note, this needs to be run before init_table()
def init_vars():
    global INIT_VARS
    global inventory,SPECIES_COL, COMMON_COL, DBH_COL, HEIGHT_COL, SPREAD_COL, COND_COL, LOC_COL, SPACE_COL
    global wb, s, runs
    try:
        obj = open(path + 'src\\config.txt', "r")
        inventory = obj.readline()
        inventory = inventory.rstrip('\n')
        wb = open_workbook(path + inventory) #READ
        s = wb.sheet_by_index(0)             #READ
        book = Workbook()                    #WRITE  
        sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)   #WRITE
        SPECIES_COL = int(obj.readline()) #27
        COMMON_COL  = int(obj.readline()) #26
        DBH_COL     = int(obj.readline()) #29
        HEIGHT_COL  = int(obj.readline()) #22 total BS
        SPREAD_COL  = int(obj.readline()) #32 also total BS
        COND_COL    = int(obj.readline()) #36 now a string!!!
        LOC_COL     = int(obj.readline()) #37 also a string!!
        SPACE_COL   = int(obj.readline()) #30 
        obj.close()
        INIT_VARS = True  
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
    setCol = raw_input("What column is the GROWSPACE DESCRIPTOR? ")
    obj.write(ctoi(setCol)+'\n')
    #init_vars()





    
       

#called in isValidEntry() since that is called before each projection,
#thereby resetting the fields
def setFields(row):
    global loc, species, common, dbh, height, spread, cond, space
    try:
        species = getCell(row, SPECIES_COL)
        common  = getCell(row, COMMON_COL)
        dbh     = float(getCell(row, DBH_COL))
        height  = float(getCell(row, HEIGHT_COL))
        spread  = float(getCell(row, SPREAD_COL))
        cond    = getCell(row, COND_COL)
        loc     = getCell(row, LOC_COL)
        space   = getCell(row, SPACE_COL)
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
        space   = getCell(row, SPACE_COL)
    


    
    
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
    if( cond  == empty_cell.value ):
        return False
    if( loc  == empty_cell.value):
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
    
#helper function to index dbh class, called in getDbhClassfromDBH()
def incofFive(num):
    mod = num % 5
    if(mod == 0):
        return num
    elif( mod >= 3):
        return num + (5 - mod)
    else:
        return num - mod

#Called by project_row() returning the appropriate* index for growth increment class
def getGrowthClassFromCond(spread, height, cond, loc, space):
    global prinfo, loop, dbh
    #BY INDEX of column vs growth increments per inch:
    # 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9  | 10 | 11 | 12
    #---|---|---|---|---|---|---|---|---|----|----|----|----
    # 1 |1.5| 2 |2.5| 3 | 4 | 5 |7.5|10 |12.5| 15 |17.5| 20
    index = 0
    #print space + " -> " + str(spaceValue(space))
    index += spaceValue(space)
    #MAX possible 4
    if(cond == "Good"):
        index += 0
    elif(cond == "Fair" ):
        index += 2
    elif(cond == "Poor" ):
        index += 3
    elif(cond == "Dead"):
        index += 4
    else:
        index += random.randint(0,4)
    #MAX possible 8
    if(loc == "Good"):
        index += 0
    elif(loc == "Fair" ):
        index += random.randint(1,3)
    else:
        index += 4
    #MAX possible 12
    if(prinfo):
        print str(index + loop) + " at " + str(dbh) 
    if(index <= 12 and loop < 1 ):
        return index
    else:
        return 12

#must be changed for each new environment
def spaceValue(growspace):
    if(growspace == "Lawn" ):
        return 1
    elif(growspace == ">4'"):
        return 2
    elif(growspace == "<4'"):
        return 3
    elif(growspace == "Sidewalk"):
        return 4
    else:
        return random.randint(0,4) #keep for a nieve even distro

#used after a set_vars() call to re-read the fields a user has passed in
#reset() is called separately to alleviate risk of a simultaneous reading and writing to config.txt 
def reset():
    init_vars()
    init_table()

#Test function, called from terminal 
def test():
    project_row(1, 10)

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
SPACE_COL   = -23

#(above and below) python vars for working with fields  
  
species = ""
common = ""
dbh    = -1
height = -1
spread = -1
cond   = -1
loc    = -1
space  = "a"

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
sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)

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
        init_vars()
    elif(cmd == 'reset'):
        reset()
    elif(cmd == 'setup'):
        set_vars()
    else:
        print "USAGE: 'grow' -- grow the inventory by some number of years"
        print "       'quit' or 'q' -- exit the program"
        print "       'setup' -- prompt to reset path and columns for inventory"
        print "       'reset' -- restablish connection with growth table and inventory, use after setup"
        
