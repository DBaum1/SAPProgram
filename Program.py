#!/usr/bin/python
#!python3
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import openpyxl
import os
from configparser import ConfigParser, NoSectionError, DuplicateSectionError
from configparser import NoOptionError
#constants
COLOR = '#2c3766'
TEXT_COLOR = '#ffffff'
DEFAULT_SIZE = '530x330'
MIN_WIDTH = 500
MIN_HEIGHT = 320
MAX_WIDTH = 570
MAX_HEIGHT = 350
LABEL_WIDTH = 55
COL_NUM = 2 #column for grid layout
#Listings to check SAP for
LISTINGS = ["Contract Number", "Contract Name", "Vendor name", "OA Amount",
            "OA Net", "OA Remaining", "Validity Start Date", "Expiration Date"]
ENTRY_LIST = [] #Store references to grid entries
DEFAULT_COLS = ["A", "B", "C", "F", "G", "", "J", "K"]

#Write user entered data to config file
def write_to_config(parent):    
    config = ConfigParser()
    try:
        config.read('config.ini')
        with open('config.ini', 'w') as f:
            for i in range(len(ENTRY_LIST)):
                val = ENTRY_LIST[i].get()
                try:
                    config.set('main', LISTINGS[i], val)
                except (NoSectionError, DuplicateSectionError):
                    init_config()
                    return None
            config.write(f)
            f.close()
    except IOError:
        init_config()
    
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
    with open('config.ini', 'w') as f:
        for i in range(len(LISTINGS)):
            config.set('main', LISTINGS[i], DEFAULT_COLS[i])
        config.write(f)
        f.close()
        
#Fills entries with values from config file
def fillGrid(parent):
    for i in range(len(LISTINGS)):
        curr_label = Label(parent, text=LISTINGS[i])
        curr_label.grid(row=i+1, column=COL_NUM, padx=5,pady=5)
        #initialize entry fields
        curr_entry = Entry(parent)
        ENTRY_LIST.append(curr_entry)
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
def showColTable():
    table = Tk()
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
    fillGrid(table)
    table.protocol("WM_DELETE_WINDOW", lambda:(write_to_config(table),
                                               table.destroy()))
    table.mainloop()
    
#Read spreadsheet sheet
def readSheet(sheet):
    num_rows = sheet.max_row
    num_cols = sheet.max_column
    print(num_rows)
    print(num_cols)
    for r in range(1, num_rows, 1):
        for c in range(1, num_cols, 1):
            print(sheet.cell(row=r, column=c).value)
            
#debugging - make sure Excel file contents is being read properly
#printing it to console
def excelDebug():
    file_path = path_label.cget("text")
    if(file_path is not None):
        print(file_path)
        wb = openpyxl.load_workbook(file_path)
        #get sheet user entered
        sheetname = sheet_entry.get()
        sheets = wb.sheetnames
        try:
            sheet = wb[sheetname]
            #readSheet(sheet)
        #Sheet doesn't exist
        except KeyError:
           messagebox.showerror("Sheet not found!", "Check sheet entry"
                                + " field for spelling,"
                                + " spacing, and capitalization."
                                + " It must exactly match the Excel doc"
                                + " sheet name.")
           
#Automate control of mouse
def importData():
    #coords = pyautogui.locateOnScreen('C:/Users/it4892/Desktop/Capture.png')
    #if(coords is None):
    #    print("not found")
    excelDebug()
    print("importData clicked")

#Excel file selection dialog
def showFileChooser(arg=None):
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

label = Label(root, wraplength=300, text="Program to transfer data "
          + "from SAP database to Excel spreaadsheet. Please log into SAP "
          + "and navigate to agreement number page before using.\n"
          + "\nHit CTRL+C to terminate program (will necessitate restarting"
              + " to continue)")
label.configure(background=COLOR)
label.configure(foreground=TEXT_COLOR)
label.pack()

#Label to show user the destination path they chose
path_label = Label(root, text="",  width=LABEL_WIDTH)
path_label.pack(pady=10)

#File Chooser Button
file_button = Button(root, text="Select destination file",
                     command=showFileChooser)
file_button.pack(pady=10)

#Sheet
sheet_label = Label(root, text="Enter Excel sheet name ex: Services "
                    + "(case sensitive)")
sheet_label.pack(pady=5)
sheet_entry = Entry()
sheet_entry.pack(pady=10)

#Column Info
col_info_btn = Button(root, text="Enter/Verify Column Information",
                      command=showColTable)
col_info_btn.pack(pady=10)

#Import Button, initially disabled
import_btn = Button(root, text="Import from SAP", state=DISABLED,
                    command=importData)
import_btn.pack()

root.mainloop()
