#!/usr/bin/python

import xlrd
import xlwt 

import datetime

from itertools import groupby
from operator import itemgetter

filepath = 'C:\Python27\'
inputFile = 'TestPython.xls'
outputFile = 'ExpenseSummary.xls'

col_width = 256 * 20             # 20 characters wide

ezxf = xlwt.easyxf
heading_xf = ezxf('font: bold on; align: wrap on, vert centre, horiz center; pattern: pattern solid, fore-colour grey25')
##color_xf = ezxf('pattern: pattern solid, fore-colour ice_blue')

style = xlwt.XFStyle()
style.num_format_str = '#,##0.00'

workbook = xlrd.open_workbook(filepath + inputFile)

#wbx = copy(workbook) # Initiating write object

Testworksheet = workbook.sheet_by_name('Test')

# ----- Start of Sorting -----
    # Sorting of any column, just give the column number to target_column variable
target_column = 0 # Sort on Date field

data = [Testworksheet.row_values(i) for i in xrange(Testworksheet.nrows)] # returns list with rows with a list of columns

labels = data[0]  # Header row

data = data[1:]  # Complete data except the header

# sort the data based on Date field
    ##data.sort(key=lambda x: x[target_column])

data.sort(key = itemgetter(target_column))

# ----- End of Sorting -----

# ----- Calculate of Income -----
L_Income = []
D_Income = {}

D_SumOfUniqueTags = {}

for tag in data:
    date_tuple = xlrd.xldate_as_tuple(tag[0],workbook.datemode)
    now = datetime.datetime(date_tuple[0] , date_tuple[1], date_tuple[2])

    if tag[3] == 'Income':
        L_Income.append((now.month, tag[2]))

    else:

        L_Rowtags = tag[3].split(',')
        L_Rowtags = [catg.strip() for catg in L_Rowtags] # Trimming spaces after split done above
          
        # -------- Functionality 1 :: Total of each unique Tag for entire input not based on month or date. ---------------------

        for newTag in L_Rowtags:
            
            if D_SumOfUniqueTags.has_key((now.month, newTag)):     # if tags are present in dictionary sum the price

                D_SumOfUniqueTags[(now.month, newTag)] = D_SumOfUniqueTags[(now.month, newTag)] + tag[2]

            else:   # if new tags then add new tag to dictionary and

                D_SumOfUniqueTags[(now.month, newTag)] = tag[2]
 	# -------- End of Functionality 1 ---------------------------------------------------------------------------------------      

groups = groupby(L_Income, itemgetter(0))
for key,value in groups:
    s = sum([ item[1] for item in value ])
    D_Income[key] = s  # store the monthly income into dictionary to match later

# ----- End of Calculate Income -----


# ------- Group the sorted output by date ---------

groups = groupby(data, itemgetter(target_column))

# ------- End of Grouping ---------

# ------- Calculate Date wise Total --------
T_DailyTotal = () # Tuple to store pair of date and total expense
L_DailyTotal = []

for key,value in groups:
    s = sum([ item[2] for item in value ]) # item[2] is Rate field
        
    date_tuple = xlrd.xldate_as_tuple(item[0],workbook.datemode)
    now = datetime.datetime(date_tuple[0] , date_tuple[1], date_tuple[2])
    
    T_DailyTotal = item[0] , s , now.month # item[0] is the date field, s is the daily total, month for this date
    L_DailyTotal.append(T_DailyTotal)

# ------- End of Calculate Date wise Total --------

# ------- Calculate Month wise Total --------

groups = groupby(L_DailyTotal, itemgetter(2))

#w_sheet = wbx.get_sheet(0)  #the sheet to write to within the writable copy

wb = xlwt.Workbook()
ws = wb.add_sheet('SummaryExpense')

month_headings = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
rowx = 1
for colx, value in enumerate(month_headings):
    ws.write(rowx, colx+1, value, heading_xf)

ws.col(0).width = col_width # Column Width
ws.write(2,0,'INCOME')
ws.write(3,0,'TOTAL EXPN')
ws.write(4,0,'SAVINGS')

for key,value in groups:
    s = sum([ item[1] for item in value ]) # item[1] is Rate field

    varMonth = key
    varExpense = (s-D_Income[key])

    ws.write(2,key,D_Income[key], style)   # Income
    ws.write(3,key,varExpense, style)     # TotalExpense
    ws.write(4,key,(D_Income[key]- varExpense), style)     # Savings
    print varMonth , "th month Income =" , D_Income[key], "Expense =" , varExpense, "Savings =" ,(D_Income[key]- varExpense) # income to be minus from expense

# ------- End of Calculate Month wise Total --------    

DupTag = [tags for mnth,tags in D_SumOfUniqueTags.keys()]
L_UniqueTag = list(set(DupTag)) # removes duplicate without considering Order.

row = 6
col = 0
for Category in L_UniqueTag:  # Print category list into Excel
    ws.write(row,col,Category)
    
    for mnth,tags in D_SumOfUniqueTags.keys():
        if Category == tags:
            ws.write(row,col+mnth,D_SumOfUniqueTags[(mnth, tags)], style)     #Print values against category in appropriate month
            
    row = row+2

ws.panes_frozen = True
ws.horz_split_pos = 5
ws.vert_split_pos = 1

wb.save(filepath + outputFile)


