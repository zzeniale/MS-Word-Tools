"""
Colour table cells in MS Word based on numerical value.

Table needs to be in its own document that contains only 1 table.
"""

import os
import docx
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import tkinter as tk
from tkinter import filedialog
import ctypes

application_window = tk.Tk()

# Ask the user to select a document
doc_path = filedialog.askopenfilename(parent=application_window,
                                    initialdir=os.getcwd(),
                                    title="Please select a file:",
                                    filetypes=[('docx files','.docx')])
application_window.destroy()

doc = docx.Document(doc_path)

#read existing table and colour cells
for row in doc.tables[0].rows:
    for cell in row.cells:
        if cell.text == "Index Score":
            continue 
        # if cell is 0
        if float(cell.text) == 0: 
            green = parse_xml (r'<w:shd {} w:fill="A0DA95"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(green)
        # if cell less than 0
        elif float(cell.text) < 0: 
            orange = parse_xml (r'<w:shd {} w:fill="FFC000"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(orange)
        # if cell greater than 0
        else: 
            blue = parse_xml (r'<w:shd {} w:fill="8ADAFF"/>'.format(nsdecls('w')))
            cell._tc.get_or_add_tcPr().append(blue)
        
doc.save(doc_path)

ctypes.windll.user32.MessageBoxW(0, "All done!", "Message", 1)