from __future__ import print_function
from os.path import join, dirname, abspath
import xlrd

#fname = join(dirname(dirname(abspath(__file__))), 'test_data', 'Cad Data Mar 2014.xlsx')
fname='WL_gn_88_nrn.xls'
# Open the workbook
xl_workbook = xlrd.open_workbook(fname)

# List sheet names, and pull a sheet by name
#
sheet_names = xl_workbook.sheet_names()
print('Sheet Names', sheet_names)

xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

# Or grab the first sheet by index 
#  (sheets are zero-indexed)
#
xl_sheet = xl_workbook.sheet_by_index(0)
print ('Sheet name: %s' % xl_sheet.name)

# Pull the first row by index
#  (rows/columns are also zero-indexed)
#
row = xl_sheet.row(0)  # 1st row

# Print 1st row values and types
#
from xlrd.sheet import ctype_text   
# Print all values, iterating through rows and columns
#
num_cols = xl_sheet.ncols   # Number of columns
	
"""
f=open('data.txt','w')
for i in range(0,10):
	f.write(str(xl_sheet.row(i)))
	f.write('\n')
f.close()
f=open('data.txt','w')
for row_idx in range(0, xl_sheet.nrows):    # Iterate through rows
    print ('-'*40)
    print ('Row: %s' % row_idx)   # Print row number
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))
        f.write(str(cell_obj))
        f.write('|')
f.write('\n')
f.close()	
"""
print xl_sheet.nrows