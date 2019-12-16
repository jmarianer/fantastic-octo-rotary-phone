# -*- coding: utf-8 -*-
"""
Created on Wed Dec 11 11:43:25 2019
This program will read a file in from TreeSize and determine what duplicates
exist and mark them "Duplicate" or "Different" in a new column.
Program inserts a new column before column A.
Comparison is of adjacent rows using columns B, D, E, and F.
Sort by column A prior to running is currently required.
@author: ad414d
"""

#Import function to read xlsx file.
from openpyxl import load_workbook

#Define variable for the TreeSize workbook.
wb=load_workbook(r'C:\\aaTanker\TreeSize\jrl me folder - duplicate.xlsx')

#Define variable for the worksheet.
sheet = wb['Custom Search']

#Insert a new column A.
sheet.insert_cols(1)
sheet.column_dimensions['A'].width=10

#Compare a row to the next.
#CR = current row, NR = next row.
for i in range(5,sheet.max_row):
    j=i+1

    CR1B=sheet.cell(row=i,column=2)
    CR1D=sheet.cell(row=i,column=4)
    CR1E=sheet.cell(row=i,column=5)
    CR1F=sheet.cell(row=i,column=6)
    NR2B=sheet.cell(row=j,column=2)
    NR2D=sheet.cell(row=j,column=4)
    NR2E=sheet.cell(row=j,column=5)
    NR2F=sheet.cell(row=j,column=6)

    Current_Row=(CR1B.value, CR1D.value, CR1E.value, CR1F.value)
    Next_Row=(NR2B.value, NR2D.value, NR2E.value, NR2F.value)

    if Current_Row == Next_Row:
        sheet.cell(row=j,column=1).value="Duplicate"
    else:
        sheet.cell(row=j,column=1).value="Different"

wb.save('C:\\aaTanker\TreeSize\jrl me folder - new.xlsx')