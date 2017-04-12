#! python3
"""
excel2text.py - writes each cell in a spreadsheet to a text file. One text file per column in the spreadsheet
"""

import openpyxl

# load file into workbook
filename = input('Enter name of file:  ')
try:
    wb = openpyxl.load_workbook(filename)
except Exception as err:
    print('something went wrong loading the workbook')
    raise

ws = wb.active

# Loop though columns in sheet, creating a text file for each column
for n, column in enumerate(ws.iter_cols()):
    text_file = open('./text_docs/excel2text_output_{}.txt'.format(n), 'w')
    # Loop through each cell in column, writing line to text file.
    for cell in column:
        if cell.value is None:  # avoids exception from .write() when cell.value is None.
            continue
        else:
            text_file.write(cell.value)
    text_file.close()

print('Done.')
