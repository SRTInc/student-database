#import  xlwt
import  xlrd
from xlutils.copy import copy
#writing the dsta in excel
rb=xlrd.open_workbook('.xls')
wb=copy(rb)
w_sheet=wb.get_sheet(0)
for i in range(8):
    a=input()
    w_sheet.write(0,i,a)
wb.save('pyexcel.xls')

print(''.join(map(str, info)))  # for printing the list without square brackets and comma


main = 34
workbook = xlrd.open_workbook("studbase.xls")
sheet = workbook.sheet_by_index(0)
wrongValue = sheet.cell_value(main, 16)
print(wrongValue)
workbook_datemode = workbook.datemode
print(workbook_datemode)
y, m, d, h, m1, s = xlrd.xldate_as_tuple(wrongValue, workbook_datemode)
print(m)

#reading the data in the excel sheet

import xlrd

# Give the location of the file
loc = ("path of file")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
sheet.cell_value(0, 0)
