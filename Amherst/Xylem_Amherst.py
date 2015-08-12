#Tree Data Projector Version 0.1 - Gabe Markarian - Last Edit 2/10/14
from xlrd import open_workbook, empty_cell
from xlwt import Workbook
import os
print "How many years ahead would you like to project?"
years = input("Enter a number: ")

#working directory
path = os.getcwd() + '\\';

##Function that uses the enclosed text file to determine the growth
##rate of a specific tree (given as parameter)
def growth_rate(tree):
    #text file with all the growth information
    data  = open(path + 'tree_rates.txt', 'r')
    for line in data:
        c = line.split(':')
        if(c[0] == tree):
            data.close
            return float(c[1])

    data.close            
    return 1
##Function that determines the weight of the deadwood in the tree on the dbh
def deadwood(state):
    if(state == "<25%"):
        return 1.0
    elif(state == "25-50%"):
        return 0.9
    elif(state == "50-75%"):
        return 0.85
    elif(state == ">75%"):
         return 0.75
    #default
    else : return 1.0


##Function that determines the weight of the trunk cond. on the dbh
def trunkcond(state):
    if(state == "Poor"):
         return 0.75
    #default, for most states
    else : return 1

def printprogress(line):
    if((line % 1000) == 0):
        print "Projecting Tree #",line

def projectDBH(row, col):
    DBHx = growth_rate(s.cell(row, 26).value)
    DEADWOODx = deadwood(s.cell(row, 31).value)
    TRUNKx = trunkcond(s.cell(row, 36).value)
    if(col == 29):
        sheet1.write(row, col, (s.cell(row, col).value) +
                                 (years*DBHx)*DEADWOODx*TRUNKx)
    else:
        sheet1.write(row, col, s.cell(row, col).value)    


def reprint(row, col):
    
        
    sheet1.write(row, col, s.cell(row, col).value)

#opening the excel workboook
wb = open_workbook(path + 'TreeData.xls')
s = wb.sheet_by_index(0)


#creating the working copy of the 
book = Workbook()
sheet1 = book.add_sheet('Sheet 1')
usable = 0
print "Projecting Tree # 0000"
##Loops that go through the excel spreadsheet and apply the above function
for row in range(s.nrows):
    project = True
    if(row == 0): continue;
    if(s.cell(row, 29).value == empty_cell.value):
        project = False
    else:
        usable += 1
    printprogress(row)
    for col in range(s.ncols):
        if(project):
            projectDBH(row, col)
        else:
            reprint(row, col)
        
for topcol in range(s.ncols):
    sheet1.write(0, topcol, s.cell(0, topcol).value)

    y = str(years);

book.save(path + y + 'Year_ProjectedTreeData.xls')

print 'The process is done! \nA new spreadsheet called ProjectedTreeData should appear in the folder.\n\n'
print str(usable) + ' Usable entries'
EOF = raw_input("Press Enter to Exit")




