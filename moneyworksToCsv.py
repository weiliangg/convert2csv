from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import re

# #workbook object 
# wb = Workbook()

#load existing spreadsheet
wb = load_workbook('testData.xlsx')

#create active worksheet 
ws = wb.active

#Grab column A
# column_a = ws['A']
# for cell in column_a:
#     if (cell.value is not None and cell.value !=''):
#         print(re.sub("[.]","", cell.value))

# column_b = ws['B']
# for cell in column_b:
#     if (cell.value is not None and cell.value !=''):
#         print(cell.value)

