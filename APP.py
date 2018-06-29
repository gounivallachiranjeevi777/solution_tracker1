import xlrd  #import excel reader
import os
from tkinter import ttk
from tkinter import *

z = os.getcwd()
workbook = xlrd.open_workbook(z+'\\Output.xlsx')
sheets = workbook.sheet_names()
required_data = []
sop = Tk()
v = Label(sop, text="FAQ's")
v.pack()
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data.append((row_values[0]))
required_data1 = list(filter(None, required_data))
def on_keyrelease(event):

    # get text from entry
    value = event.widget.get()
    value = value.strip().lower()

    # get data from required_data1
    if value == '':
        data = required_data1
    else:
        data = []
        for item in required_data1:
            if value in item.lower():
                data.append(item)                

    # update data in listbox
    listbox_update(data)


def listbox_update(data):
    # delete previous data
    listbox.delete(0, 'end')

    # sorting data 
    data = sorted(data, key=str.lower)

    # put new data
    for item in data:
        listbox.insert('end', item)


def on_select(event):
    global ert
    ert = event.widget.get(event.widget.curselection())
    entry.delete(0,END)
    entry.insert(0,ert)

# --- main ---
entry = ttk.Entry(sop, width=100)
entry.pack()
entry.bind('<KeyRelease>', on_keyrelease)

fa = ttk.Scrollbar(sop)
listbox = Listbox(sop, width=100, height = 25, yscrollcommand=fa.set)
listbox.pack(side = LEFT)
fa.pack(side = LEFT, fill=Y)
fa.config(command = listbox.yview)
listbox.bind('<<ListboxSelect>>', on_select)
listbox_update(required_data1)


def z():
    y.delete('1.0',END)

def az():
    exe = entry.get()
    for sheet in workbook.sheets():
            for rowidx in range(sheet.nrows):
                row = sheet.row(rowidx)
                for colidx, cell in enumerate(row):
                    if cell.value == exe :
                        aov = sheet.cell_value(rowidx,colidx+1)
                        y.insert(END,aov+"\n")

s = ttk.Button(sop, text="SUBMIT", command = az)
s.pack()
ww = ttk.Button(sop, text="CLEAR",command = z)
ww.pack()
f = ttk.Scrollbar(sop)
y = Text(sop, yscrollcommand=f.set)
y.pack(side = LEFT)
f.pack(side = RIGHT, fill=Y)
f.config(command = y.yview)
sop.mainloop()

