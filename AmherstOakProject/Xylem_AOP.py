#Xylem open source tree projection for Tree Wardens meeting c/o Rick Harper
from xlrd import open_workbook, empty_cell
from xlwt import Workbook
import os
import time

#get working directory
path = os.getcwd() + '\\'

#starts a chain of method calls that project each entry in an inventory
def project(years):
    global INVALIDS #provide information about execution at the end
    print_row(0, True)
    for r in range(1, s.nrows):
        project_row(r, years)
    timestamp = time.strftime('%Y%m%d%H%M%S')
    book.save(path + str(years) + 'Projected'+timestamp+'.xls')
    print "There were " + str(INVALIDS) + " unusable entries out of " + str(s.nrows)

#uses current dbh and rate values to extrapolate a new dbh for a given number of years
#years is gotten originally from command line and passed down through project()
def project_row(row, years):
    global dbh
    if(isValidEntry(row)):
        for name in range(len(RATES)):
            if(common == RATES[name][0]):
                dbh = dbh + RATES[name][1]*years
                #currently without a failsafe
        print_row(row, True)
    else:
        print_row(row, False)

#communicates new dbh values and tree name info to the projected inventory
def print_row(row, valid):
    global WRITE_ROW, INVALIDS
    if(row == 0):  #special case where column headers need to be printed
        sheet1.write(0, SPECIES_COL, "Species")
        sheet1.write(0, COMMON_COL,  "Common" )
        sheet1.write(0, DBH_COL,     "DBH"    )
        return 
        
    #print common name, dbh, anything else?
    if(valid):
        #made the choice to only reprint necessary fields since thats all iTree uses
        sheet1.write(WRITE_ROW, SPECIES_COL, species)
        sheet1.write(WRITE_ROW, COMMON_COL,  common )
        sheet1.write(WRITE_ROW, DBH_COL,     dbh    )
        WRITE_ROW += 1
    else:
        INVALIDS += 1
    
    
#checks if a row has the necessary criteria for projection
#also sets the fields so individual entries can be read, written, and manipulated
def isValidEntry(row):
    if(row > s.nrows):
        return False
    setFields(row)
    if(species == empty_cell.value):
        return False
    if(common  == empty_cell.value):
        return False
    #may be changed later to handle all dbh's
    if(dbh  == empty_cell.value or dbh == 0 or type(dbh) is not float):
        return False
    return True

#sets local variables to those read from inventories, exclusively called by validity test above
def setFields(row):
    global species, common, dbh
    try:
        species = getCell(row, SPECIES_COL)
        common  = getCell(row, COMMON_COL)
        dbh     = float(getCell(row, DBH_COL))
    except ValueError:
        #On excpetion: let isValid entry decide validity of raw
        #cell value, trivially it will return false
        species = getCell(row, SPECIES_COL)
        common  = getCell(row, COMMON_COL)
        dbh     = getCell(row, DBH_COL)

#responsible for initializing values in the rates table
#called just before the REPL starts up
def init_all():
    RATES[0].append("Oak, Swamp White")
    RATES[0].append( 0.52 )
    RATES[1].append( "Oak, Red" )
    RATES[1].append( 0.37 )
    
    
#easy access to cell vals
def getCell(r, c):
    return s.cell(r, c).value

def test():
    global INVALIDS
    print_row(0, True)
    for i in range(1, 7):
        print_row(i, isValidEntry(i))
    print INVALIDS
    book.save(path + "testerproj.xls")
#oak, swamp white:0.52
#Red Oak:0.37
SPECIES_COL = 2
species     = ""

COMMON_COL  = 3
common      = ""

DBH_COL     = 6
dbh         = -1

RATES = [[] for i in range(2)]

#READ FROM
wb = open_workbook(path + "HCMP092714.xls")
s  = wb.sheet_by_index(0)

#WRITE TO
WRITE_ROW = 1;  #tells print functions where to print in projected versions
INVALIDS = 0; #tracks number of invalid entries
book = Workbook()
sheet1 = book.add_sheet('ProjectedData')

init_all()
print "all set, type any non-command for usage"
#######################################
#REPL code, will run until termination#
#######################################
cmd = ""

while(cmd != 'quit' and cmd != 'q'):
    cmd = raw_input("> ")
    if(cmd == 'test'):
        test()
    elif(cmd == 'grow' or cmd == 'project'):
        print " How many years of growth?"
        years = input(">> ")
        project(years)
    elif(cmd == 'q' or cmd == "quit"):
        print "Have a good day!"
    else:
        print "USAGE: `test` run your test method\n`grow` run a projection"
