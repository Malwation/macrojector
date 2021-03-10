import os, sys
from os import path
import platform
import signal
import win32com.client
import winreg

objExcel = win32com.client.Dispatch("Excel.Application")
objExcel.Visible = False 
version = objExcel.Application.Version
objExcel.Application.Quit()
del objExcel

def enableVbom():
    keyval = "Software\\Microsoft\Office\\"  + version + "\\Excel\\Security"
    Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
    winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,1) 
    winreg.CloseKey(Registrykey)
        
    
def disableVbom():
    keyval = "Software\\Microsoft\Office\\"  + version + "\\Excel\\Security"
    Registrykey = winreg.CreateKey(winreg.HKEY_CURRENT_USER,keyval)
    winreg.SetValueEx(Registrykey,"AccessVBOM",0,winreg.REG_DWORD,0) 
    winreg.CloseKey(Registrykey)

def excelMacro(macropath,output):
    enableVbom()
    macro = open(macropath, "r").read()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Add()
    excelModule = workbook.VBProject.VBComponents("ThisWorkbook")
    excelModule.CodeModule.AddFromString(macro)
    excel.DisplayAlerts=False
    xlRDIAll = 99
    workbook.RemoveDocumentInformation(xlRDIAll)
    workbook.SaveAs(output, FileFormat=52)
    excel.Workbooks(1).Close(SaveChanges=1)
    excel.Application.Quit()
    del excel
    disableVbom()

macropath = sys.argv[1]
excelpath = sys.argv[2]

excelMacro(macropath,excelpath)