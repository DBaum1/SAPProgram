#!/usr/bin/python
#!python3
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import openpyxl
import os
from configparser import ConfigParser, NoSectionError, DuplicateSectionError
from configparser import NoOptionError
import time
from shutil import copyfile

#constants
COLOR = '#2c3766'
TEXT_COLOR = '#ffffff'
DEFAULT_SIZE = '530x380'
MIN_WIDTH = 500
MIN_HEIGHT = 365
MAX_WIDTH = 570
MAX_HEIGHT = 390
LABEL_WIDTH = 55
COL_NUM = 2 #column for grid layout
#Listings to check SAP for
LISTINGS = ["Contract Number", "Contract Name", "Vendor name", "OA Amount",
            "OA Net", "OA Remaining", "Validity Start Date", "Expiration Date"]
ENTRY_LIST = [None] * len(LISTINGS) #Store references to grid entries
ENTRY_VALS = [None] * len(LISTINGS) #Store values of grid entries
DEFAULT_COLS = ["A", "B", "C", "F", "G", "", "J", "K"]
PATH = 'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'

#transfers data from SAP fields to excel file
def sap_transfer():
    ############################
    #testing functionality
    app = Application(backend="win32").start(PATH)
    print("started app")
    # describe the window inside saplogon.exe process
    dlg_spec = app.SAP_Logon
    print("described window")
    actionable_dlg = dlg_spec.wait('visible')
    print("logon visible")
    field = dlg_spec['Edit0']
    field.type_keys('hello')
    print("keys typed")
    dlg_spec['Variable Logon...'].click()
    var_window = app.Logon_To_System
    var_window['Cancel'].close_click()
    print("cancel clicked")
    field.type_keys('again')
    """
        #app = Application(backend="uia").connect(path=PATH)
        Properties = Desktop(backend='win32').SAP_Logon
        #Type hello into entry field
        field = Properties['Edit0']
        field.type_keys('hello')
        #Click variable logon
        Properties['Button2'].click()
        Properties2 = Desktop(backend='win32').Logon_to_System
        #Exit variable logon
        Properties2['Cancel'].close_click()
        #get text from field
    """
    #############################
    
#copies original file, timestamps backup
def save_backup():
    src = path_label.cget("text")
    components = os.path.splitext(src)
    root = components[0]
    ext = components[1]
    time_tuple = time.localtime()
    format_time = time.strftime('_%m_%d_%Y_%Hh_%Mm_%Ss', time_tuple)
    root_format = root + format_time
    dest = root_format + ext
    copyfile(src, dest)

#Write user entered data to config file
def write_to_config(parent):    
    config = ConfigParser()
    try:
        config.read('config.ini')
        with open('config.ini', 'w') as f:
            for i in range(len(ENTRY_LIST)):
                val = ENTRY_LIST[i].get()
                ENTRY_VALS[i] = val
                try:
                    config.set('main', LISTINGS[i], val)
                except (NoSectionError, DuplicateSectionError):
                    init_config()
                    return None
            config.write(f)
            f.close()
    except IOError:
        init_config()
    col_info_btn.configure(state=NORMAL)

    
#Reads from config
def read_from_config(parent):
    config = ConfigParser()
    config.read('config.ini')
    #loop through entry fields
    for i in range(len(ENTRY_LIST)):
        curr_entry = ENTRY_LIST[i]
        try:
            val = config.get('main', LISTINGS[i])
            curr_entry.insert(0, val)
            ENTRY_VALS[i] = val
        #Config corrupted, remake
        except (NoSectionError, DuplicateSectionError):
            init_config()
        except NoOptionError:
            pass
        
#Creates config file
def init_config():
    config = ConfigParser()
    config.write('config.ini')
    config.add_section('main')
    config.set('main', 'path', PATH)
    with open('config.ini', 'w') as f:
        for i in range(len(LISTINGS)):
            config.set('main', LISTINGS[i], DEFAULT_COLS[i])
        config.write(f)
        f.close()
        
#Fills entries with values from config file
def fill_grid(parent):
    for i in range(len(LISTINGS)):
        curr_label = Label(parent, text=LISTINGS[i])
        curr_label.grid(row=i+1, column=COL_NUM, padx=5,pady=5)
        #initialize entry fields
        curr_entry = Entry(parent)
        ENTRY_LIST[i] = curr_entry
        curr_entry.grid(row=i+1, column=COL_NUM + 1, padx=5,pady=5)
    #config file exists 
    try:
        f = open('config.ini', 'r')
    #create config from scratch
    except IOError:
        init_config()
    #load in values from config file
    read_from_config(parent)
    
#Shows table where user inputs what information is in each column
#Filled in by default. Blank spaces are skipped
def show_col_table():
    col_info_btn.configure(state=DISABLED)
    table = Toplevel()
    table.title("Set Column Info")
    table.geometry(DEFAULT_SIZE)
    col_table_label = Label(table, wraplength = 500, text="Enter the column"
                            +" letter that the relevant information "
                            + " can be found in (ex: A, B, AA etc.)"
                            + " If a column's value is calculated with"
                            + " a function, leave that field blank")
    col_table_label.grid(columnspan = 8)
    table.minsize(width=MIN_WIDTH, height = MIN_HEIGHT)
    table.maxsize(width=MAX_WIDTH, height = MAX_HEIGHT)
    fill_grid(table)
    table.protocol("WM_DELETE_WINDOW", lambda:(write_to_config(table),
                                               table.destroy()))
    
#Read spreadsheet sheet
def read_sheet(sheet):
    num_rows = sheet.max_row
    num_cols = sheet.max_column
    print(num_rows)
    print(num_cols)
    for r in range(1, num_rows, 1):
        for c in range(1, num_cols, 1):
            print(sheet.cell(row=r, column=c).value)
           
#Automate control of mouse
def import_data():
    file_path = path_label.cget("text")
    try:
        wb = openpyxl.load_workbook(file_path)
        #get sheet user entered
        sheetname = sheet_entry.get()
        sheets = wb.sheetnames
        sheet = wb[sheetname]
        save_backup()
        #read_sheet(sheet)
    #File no longer exists at path
    except IOError:
            messagebox.showerror("File not found!", "File not found"
                                 + " at selected path. Check to make"
                                 + " sure it wasn't deleted or moved.")
    #Sheet not present
    except KeyError:
           messagebox.showerror("Sheet not found!", "Check sheet entry"
                                + " field for spelling,"
                                + " spacing, and capitalization."
                                + " It must exactly match the Excel doc"
                                + " sheet name.")
    #User ends program early
    except KeyboardInterrupt:
        root.destroy()
    print("importData clicked")
        
#Excel file selection dialog
def show_file_chooser(arg=None):
    filename = askopenfilename(parent=None, title = "Select file",
                               filetypes = [(("Excel (.xlsx)","*.xlsx"))])
    length = len(filename)
    if(length > LABEL_WIDTH):
        path_label.configure(width=length)
    path_label.config(text=filename)
    #User must have selected file
    if(filename):
        file_path = filename
        import_btn.configure(state=NORMAL)
        
#GUI
root = Tk()
root.title("SAP to Excel")
root.geometry(DEFAULT_SIZE)
root.configure(background=COLOR)
root.minsize(width=MIN_WIDTH, height = MIN_HEIGHT)
root.maxsize(width=MAX_WIDTH, height = MAX_HEIGHT)

label1 = Label(root, wraplength=500, text="Program to transfer data "
          + "from SAP database to Excel spreaadsheet. Please log into SAP "
          + "and navigate to agreement number page before using.\n")
label1.configure(background=COLOR)
label1.configure(foreground=TEXT_COLOR)
label1.pack()

label2 = Label(root, wraplength=500, text=" NOTE: This program creates a"
               + " timestamped backup of any file it modifies. It is"
               + " HIGHLY recommended not to delete it until you have"
               + " verified all the new information is valid.\n")
label2.configure(background=COLOR)
label2.configure(foreground='red')
label2.pack()

label3 = Label(root, wraplength=500, text=" WARNING: Terminating the program"
               + " early via CTRL+C or clicking while the program is running"
               + " will cause problems that necessitate restarting.")
label3.configure(background=COLOR)
label3.configure(foreground=TEXT_COLOR)
label3.pack()

#Label to show user the destination path they chose
path_label = Label(root, text="",  width=LABEL_WIDTH)
path_label.pack(pady=10)

#File Chooser Button
file_button = Button(root, text="Select destination file",
                     command=show_file_chooser)
file_button.pack(pady=10)

#Sheet
sheet_label = Label(root, text="Enter Excel sheet name ex: Services "
                    + "(case sensitive)")
sheet_label.pack(pady=5)
sheet_entry = Entry()
sheet_entry.pack(pady=10)

#Column Info
col_info_btn = Button(root, text="Enter/Verify Column Information",
                      command=show_col_table)
col_info_btn.pack(pady=10)

#Import Button, initially disabled
import_btn = Button(root, text="Import from SAP", state=DISABLED,
                    command=import_data)
import_btn.pack()

root.mainloop()
