#!/usr/bin/python
#!python3
from tkinter import *
from tkinter.filedialog import askopenfilename
import openpyxl
import os
from configparser import ConfigParser, NoSectionError, DuplicateSectionError,\
                     NoOptionError
import time
from shutil import copyfile
from pywinauto.application import Application

#constants
COLOR = '#063256'
FONT = 11
DEFAULT_SIZE = '560x450'
MIN_WIDTH = 550
MIN_HEIGHT = 440
MAX_WIDTH = 570
MAX_HEIGHT = 460
PATH = 'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
LISTINGS = ["Contract Number", "Contract Name", "Vendor name", "OA Amount",
            "OA Net", "OA Remaining", "Validity Start Date", "Expiration Date"]

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
    #get text from field and store in variable
    text = field.text_block()
    print(text)
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

sap_transfer()
