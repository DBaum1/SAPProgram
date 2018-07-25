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
from pywinauto import Desktop

#constants
PATH = 'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
LISTINGS = ["Contract Number", "Contract Name", "Vendor name", "OA Amount",
            "OA Net", "OA Remaining", "Validity Start Date", "Expiration Date"]

#transfers data from SAP fields to excel file
def sap_transfer():
    ############################

    #testing functionality
    Properties = Desktop(backend='win32').Display_Contract_Header_Data
    
    #app = Application().connect(path=PATH)
    container = 'Afx:60310000:1008'
    num = '4600014943'
    Properties2 = Properties[container]['AfxMDIFrame110']
    #dlg_spec = app['Display Contract:Item Overview']['Button4'].click()
    #dlg_spec = app['Display Contract:Header Data']['Afx:60310000:1008'].click()
    #dlg_spec.print_control_identifiers()
    #dlg_spec.TypeKeys('{ENTER}')
    #actionable_dlg = dlg_spec.wait('visible')
    #dlg_spec['SAP\'s Advanced Treelist'].draw_outline()
    #dlg_spec.print_control_identifiers()
    print("done")
    #############################

sap_transfer()
