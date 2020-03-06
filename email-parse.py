#!/usr/bin/env python

import functions
import glob
import os
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font
import tkinter
from tkinter import messagebox, filedialog

today = datetime.now()
# snapshots older than 'age' are considered old.
age = 30
# outlook mails are encoded as 'latin-1'
encoding = 'latin-1'
# table headers as formated in the emails)
vmTableHeaders = ['VMname', 'VMOwner', 'SSName', 'SSCreated', 'SSDescription']
# headers for the output excel table
excelHeaders = ['ServerName', '', 'ServerOwner', 'Date', 'Description']
oldSnapshots = list()
filesWithoutHTML = list()

# ask user for folder with emails (as html)
root = tkinter.Tk()
root.withdraw()
tkinter.messagebox.showinfo('Folder Select', 'Please select the folder that contains the emails (saved as html-Files)')
filePath = filedialog.askdirectory() + '/'

filesInFolder = glob.glob(os.path.join(filePath, '*.htm'))
filesInFolder.extend(glob.glob(os.path.join(filePath, '*.html')))

if not filesInFolder:
    tkinter.messagebox.showinfo \
        ('Parsing aborted', 'The selected folder doesn\'t seem to contain any htm- or html-files')
    quit()

for file in filesInFolder:
    emailTables = functions.getTablesFromHTML(file, encoding)
    vmTable = list()
    # check if one of the tables is the one with the correct header
    for table in emailTables:
        if table[0] == vmTableHeaders:
            vmTable = table
    if not vmTable:
        filesWithoutHTML.append(os.path.realpath(file))
        continue

    # test all rows except the first if date is older than 'age' and vm-name starts with 'z'
    for row in table[1:]:
        if row[vmTableHeaders.index('VMname')].startswith('z') and (today - timedelta(days=age)) > \
                datetime.strptime(row[vmTableHeaders.index('SSCreated')], '%m/%d/%Y %I:%M:%S %p'):
            oldSnapshots.append(row)

if not oldSnapshots:
    tkinter.messagebox.showinfo \
        ('Parsing aborted', 'None of the files in \"' + filePath + '\" contains a table of VM snapshots.')
    quit()

# create unique server owner list
serverOwners = []
for snapshot in oldSnapshots:
    if snapshot[vmTableHeaders.index('VMOwner')] not in serverOwners:
        serverOwners.append(snapshot[vmTableHeaders.index('VMOwner')])
serverOwners.sort()

# creating output xlsx-file
book = Workbook()
sheet = book.active
sheet.append(excelHeaders)
# append snapshot info to sheet, group them by owner
for owner in serverOwners:
    for snapshot in oldSnapshots:
        if snapshot[vmTableHeaders.index('VMOwner')] == owner:
            line = [snapshot[vmTableHeaders.index('VMname')].rstrip(),
                    '-',
                    snapshot[vmTableHeaders.index('VMOwner')].rstrip(),
                    datetime.strptime(snapshot[vmTableHeaders.index('SSCreated')].split(' ')[0], '%m/%d/%Y').strftime(
                        '%Y-%m-%d').rstrip(),
                    snapshot[vmTableHeaders.index('SSDescription')].rstrip()]
            sheet.append(line)

# formatting: first row bold and bigger; all column widths adjusted to maximum length of content
for cell in sheet["1:1"]:
    cell.font = Font(size=16, bold=True)
for column_cells in sheet.columns:
    length = max(len(str(cell.value)) for cell in column_cells)
    sheet.column_dimensions[column_cells[0].column_letter].width = length + 1

# add the filenames where no tables of snapshots were found
rowPointer = sheet.max_row + 1
if filesWithoutHTML:
    filesWithoutHTML.insert(0, 'Files in Folder without a table of VM snapshots:')
    for entry in filesWithoutHTML:
        rowPointer += 1
        sheet.cell(rowPointer, 1, entry).font = Font(color='FF0000')

errorString = 'Not executed yet.'
saveFile = filePath + 'alteSnapshots-' + today.strftime('%Y-%m-%d') + '.xlsx'
while errorString:
    try:
        book.save(saveFile)
        os.startfile(os.path.realpath(filePath))
        errorString = None
    except OSError as e:
        ok = tkinter.messagebox.askokcancel \
            ('Write Error', 'The file \"' + saveFile + '\" is open.\n' \
             'Close it and press OK to overwrite it or press cancel to abort the program.')
        if ok:
            # in case I want to do something with the error string later
            errorString = e
        else:
            errorString = None
