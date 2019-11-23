#This program will read in an Excel file and print the contents of a specific cell.

#import openpyxl module function load_workbook
from openpyxl import load_workbook

#Define variable name for the workbook
wb = load_workbook('samplebook.xlsx')

#Define variable name for worksheet
sheet = wb['Sheet1']

#Define variable for the cell to read
cell = sheet.cell(1,1)

#Define varible for cell value to print
my_val = cell.value

#Print the cell value
print(my_val)

#Read in values of cells C1 through C5 and print
for i in range(1,6):
    C_Cell = sheet.cell(i,3)
    C_Cell_value = sheet.cell(i,6)
    print(C_Cell.value)
