#!/usr/bin/python
#!python3
from tkinter import *
from tkinter.filedialog import askopenfilename
import openpyxl
from openpyxl.utils import cell
import os
from configparser import ConfigParser, NoSectionError, DuplicateSectionError,\
                     NoOptionError, MissingSectionHeaderError
import time
from shutil import copyfile
from pywinauto.application import Application
import pywinauto
import pyautogui

#constants
COLOR = '#063256'
FONT = 11
PATH = 'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
#Listings to check SAP for
LISTINGS = [('Contract Number','1'), ('Contract Name','2'),('Vendor name','3'),
        ('OA Amount', '6'), ('OA Net', '7'), ('OA Remaining', '-1'),
        ('Validity Start Date', '10'), ('Expiration Date', '11')]
ENTRY_LIST = [None] * len(LISTINGS) #Store references to grid entries

def info_import():
    print("importing")
    x = 25
    y = 160
    app = Application(backend="win32").connect(path=PATH)
    dlg_spec = app.SAP_Logon
    #dlg_spec.print_control_identifiers()
    dlg_spec['Edit3'].draw_outline()
    rect = dlg_spec['Edit3'].rectangle()
    print(dlg_spec['Edit3'].rectangle())
    print(rect.left)
    print(rect.top)

    #pywinauto.mouse.move(coords=(1154+1, 71+1))
##    pywinauto.mouse.move(coords=(x, y))
    #pywinauto.mouse.click(button='left', coords=(1154+1, 71+1))
##    pywinauto.mouse.press(button='left', coords=(x, y))
##    pywinauto.mouse.move(coords=(x+100, y))
##    pywinauto.mouse.move(coords=(x+101, y))
    #dlg_spec.SetFocus().SendKeys('^c')
    #pywinauto.mouse.click(button='left', coords=(x+50, y))


#App
info_import()
