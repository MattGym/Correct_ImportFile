# Program do porównywania db dumpa oraz bazy danych dostarczonej przez stocznię.

import openpyxl
# from openpyxl.comments import Comment
# from openpyxl.styles import PatternFill
# import pandas as pd
#tu jest zmiana z maca
import tkinter
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox

from Functions import *


class CreateToolTip(object):
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.close)

    def enter(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a toplevel window
        self.tw = tkinter.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tkinter.Label(self.tw, text=self.text, justify='left', relief='solid', borderwidth=1,
                              font=("times", "15", "normal"))
        label.pack(ipadx=1)

    def close(self, event=None):
        if self.tw:
            self.tw.destroy()


# Graphic animation function
root = Tk(className=' CdA DataBase Comparator')
root.geometry('790x185')
root.resizable(False, False)
# Variables
# File 1 Variables

file_path1 = ''
file_path_txt1 = StringVar()
file_path_txt1.set('Select import file to correct')
wb1 = openpyxl.Workbook()
active_sheet1 = wb1.active
max_rows1 = 0
max_col1 = 0

file1_col_tag = 0
file1_col_loop = 0
file1_col_package = 0
file1_col_description = 0
file1_col_min = 0
file1_col_max = 0
file1_col_unit = 0
file1_col_fbc = 0
file1_col_ibc = 0
file1_col_card = 0
file1_col_channel = 0
file1_col_instrument_code = 0
file1_col_signal_type = 0
file1_col_modbus_address = 0
file1_col_bit = 0
file1_col_gain = 0
file1_col_slave = 0
file1_col_link_signal_type = 0

# File 2 Variables
file_path2 = ''
file_path_txt2 = StringVar()
file_path_txt2.set('System - Alarm mapping table')
wb2 = openpyxl.Workbook()
active_sheet2 = wb2.active
max_rows2 = 0
max_col2 = 0


def analyze_file():
    global active_sheet1
    global max_rows1
    global max_col1
    global active_sheet2
    global max_rows2
    global max_col2

    wb1.active = wb1[sheet_choose1.get()]
    active_sheet1 = wb1.active
    max_rows1 = wb1.active.max_row
    max_col1 = wb1.active.max_column
    wb2.active = wb2[sheet_choose2.get()]
    active_sheet2 = wb2.active
    max_rows2 = wb2.active.max_row
    max_col2 = wb2.active.max_column
    fill_col_numbers()
    print(file_path1)
    print(file_path2)

    if checkbox1_var.get() == 1:
        print('DOne1')

    if checkbox2_var.get() == 1:
        print('Done2')

    if checkbox3_var.get() == 1:
        print('Done3')

    if checkbox4_var.get() == 1:
        print('Done4')

    print('saving 1st ')
    wb1.save(file_path1)
    print('saving 2nd')
    wb2.save(file_path2)
    print('Finish')


def fill_col_numbers():
    # DNA DUMP
    global file1_col_tag
    global file1_col_loop
    global file1_col_package
    global file1_col_description
    global file1_col_min
    global file1_col_max
    global file1_col_unit
    global file1_col_fbc
    global file1_col_ibc
    global file1_col_card
    global file1_col_channel
    global file1_col_instrument_code
    global file1_col_signal_type
    global file1_col_modbus_address
    global file1_col_bit
    global file1_col_gain
    global file1_col_slave
    global file1_col_link_signal_type

    file1_col_tag = get_col_no(active_sheet1, '$(TAG)', max_col1)
    file1_col_loop = get_col_no(active_sheet1, '$(LOOP)', max_col1)
    file1_col_package = get_col_no(active_sheet1, '$(PACKAGE)', max_col1)
    file1_col_description = get_col_no(active_sheet1, '$(NAME)', max_col1)
    file1_col_min = get_col_no(active_sheet1, '$(MIN)', max_col1)
    file1_col_max = get_col_no(active_sheet1, '$(MAX)', max_col1)
    file1_col_unit = get_col_no(active_sheet1, '$(UNIT)', max_col1)
    file1_col_fbc = get_col_no(active_sheet1, '$(FBC)', max_col1)
    file1_col_ibc = get_col_no(active_sheet1, '$(IBC)', max_col1)
    file1_col_card = get_col_no(active_sheet1, '$(CARD)', max_col1)
    file1_col_channel = get_col_no(active_sheet1, '$(CHANNEL)', max_col1)
    file1_col_instrument_code = get_col_no(active_sheet1, '$(INSTRUMENT_CODE)', max_col1)
    file1_col_signal_type = get_col_no(active_sheet1, '$(TEMPLATE)', max_col1)
    file1_col_modbus_address = get_col_no(active_sheet1, '$(LIS_ADDR)', max_col1)
    file1_col_bit = get_col_no(active_sheet1, '$(LIS_BIT)', max_col1)
    file1_col_gain = get_col_no(active_sheet1, '$(LIS_GAIN)', max_col1)
    file1_col_slave = get_col_no(active_sheet1, '$(LIS_SLAVE)', max_col1)
    file1_col_link_signal_type = get_col_no(active_sheet1, '$(LIS_SIGNED)', max_col1)


def choose_file1():
    global file_path1
    global wb1
    root.filename = filedialog.askopenfilename(title='Choose file to open',
                                               filetypes=(('xlsx', '*.xlsx'), ('xls', '*.xls')))
    if len(root.filename) > 0:
        file_path1 = root.filename
        file_path_txt1.set('File: ' + file_path1)
    if len(file_path1) > 0:
        wb1 = openpyxl.load_workbook(file_path1)
        sheet_names1 = [wb1.sheetnames]
        sheet_choose1['values'] = tuple(sheet_names1[0])
        sheet_choose1.current(0)
        if len(file_path2) > 2:
            button_analyze['state'] = tkinter.NORMAL


def choose_file2():
    global file_path2
    global wb2
    root.filename = filedialog.askopenfilename(title='Choose file to open',
                                               filetypes=(('xlsx', '*.xlsx'), ('xls', '*.xls')))
    if len(root.filename) > 0:
        file_path2 = root.filename
        file_path_txt2.set('File: ' + file_path2)
    if len(file_path2) > 0:
        wb2 = openpyxl.load_workbook(file_path2)
        sheet_names2 = [wb2.sheetnames]
        sheet_choose2['values'] = tuple(sheet_names2[0])
        sheet_choose2.current(0)
        if len(file_path1) > 0:
            button_analyze['state'] = tkinter.NORMAL
# ------- Graphic user interface --------
# ---------------------------------------


file_label1 = Label(root, textvariable=file_path_txt1, width=50, anchor='w', relief='groove')
file_label1.place(x=10, y=20)
sheet_choose_select1 = tkinter.StringVar()
sheet_choose1 = ttk.Combobox(root, textvariable=sheet_choose_select1, width=10, height=1)
sheet_choose1.place(x=470, y=17)
button_select1 = Button(root, text='Select', command=choose_file1, height=1, width=5)
button_select1.place(x=590, y=16)

button_select2 = Button(root, text='Select', command=choose_file2, height=1, width=5)
button_select2.place(x=590, y=46)
file_label2 = Label(root, textvariable=file_path_txt2, width=50, anchor='w', relief='groove')

file_label2.place(x=10, y=49)
sheet_choose_select2 = tkinter.StringVar()
sheet_choose2 = ttk.Combobox(root, textvariable=sheet_choose_select2, width=10, height=1)
sheet_choose2.place(x=470, y=47)

button_analyze = Button(root, text='Analyze', command=analyze_file, height=3, width=8, state=tkinter.DISABLED)
button_analyze.place(x=675, y=17)

labelframe1 = ttk.Labelframe(root, width=770, height=95, labelanchor=NW, text='Check options')
labelframe1.place(x=10, y=80)

checkbox1_var = IntVar(root, 1)
checkbox1 = Checkbutton(root, text='Update Al & Msg group', variable=checkbox1_var, onvalue=1,
                        offvalue=0, height=1)
checkbox1.place(x=20, y=105)
checkbox1_tt = CreateToolTip(checkbox1, 'Update alarm and message groups according to given table')

checkbox2_var = IntVar(root, 0)
checkbox2 = Checkbutton(root, text='Update Am100 alarm & connections', variable=checkbox2_var, onvalue=1,
                        offvalue=0, height=1)
checkbox2.place(x=20, y=127)
checkbox2_tt = CreateToolTip(checkbox2, 'Update Am100 Alarms and EXT1..4 connections with alarm prio')

checkbox3_var = IntVar(root, 0)
checkbox3 = Checkbutton(root, text='Merge Alarm and Control limits', variable=checkbox3_var, onvalue=1,
                        offvalue=0, height=1)
checkbox3.place(x=20, y=149)
checkbox3_tt = CreateToolTip(checkbox3, 'Merging alarm and control limits for Am templates if need')

checkbox4_var = IntVar(root, 0)
checkbox4 = Checkbutton(root, text='Spare', variable=checkbox4_var, onvalue=1,
                        offvalue=0, height=1)
checkbox4.place(x=260, y=105)
checkbox4_tt = CreateToolTip(checkbox4, 'New function')
root.mainloop()
