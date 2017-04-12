#! python3
"""
text2excel.py - read all text documents in a directory and write them to a spreadsheet by lines, each file located
in a separate column.

Future work: code to take directory from command line
"""

import os
import openpyxl

filepath = os.path.abspath('./text_docs/')

# Find all text files in filepath
print('Finding files...')
list_of_text_files = []
for root, dirs, files in os.walk(filepath):
    for file in files:
        path = os.path.join(root, file)
        if path.endswith('.txt'):
            list_of_text_files.append(path)

# create workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Text2Excel'

print('Reading and writing files...')
# loop through files in list reading by lines
for n, text_file in enumerate(list_of_text_files):
    fileobj = open(text_file, 'r')
    lines = fileobj.readlines()
    # for each file write each line in a column specific for that file.
    for i, line in enumerate(lines):
        ws.cell(row=i + 1, column=n + 1).value = line
    fileobj.close()

wb.save('text2excel_output.xlsx')

print('Done.')
