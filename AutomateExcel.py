import os
import time

import win32gui
import pyautogui
import xlwings

from FileOperationsReport import reportlog


def openfile(filename,filenum):
    reportlog(filename,filenum)
    wb = xlwings.Book(filename)
    reportlog("open",filenum)
    time.sleep(2)
    if win32gui.GetWindowText(1771144) == '&No':
        win32gui.SetForegroundWindow(1771144)
        time.sleep(2)
        x, y, z, w = win32gui.GetWindowRect(1771144)
        pyautogui.click(x, y)
        time.sleep(4)
    temp_Handle = win32gui.FindWindow('XLMAIN', None)
    win32gui.SetForegroundWindow(temp_Handle)


def editfile(filename,filenum):
    temp_Handle = win32gui.FindWindow('XLMAIN', None)
    win32gui.SetForegroundWindow(temp_Handle)
    #print(xlwings.Range('A1').value)
    xlwings.Range('A3').value = 'Abc'
    reportlog("edit",filenum)


def savefile(filename,filenum):
    temp_Handle = win32gui.FindWindow('XLMAIN', None)
    win32gui.SetForegroundWindow(temp_Handle)
    pyautogui.hotkey('ctrl', 's')
    reportlog("saved",filenum)


def closefile(filename,filenum):
    temp_Handle = win32gui.FindWindow('XLMAIN', None)
    win32gui.SetForegroundWindow(temp_Handle)
    pyautogui.hotkey('alt', 'f4')
    reportlog("closed",filenum)
    time.sleep(2)
    # if win32gui.GetWindowText(1378600) == "&Yes":
    #   win32gui.SetForegroundWindow(1378600)
    #   time.sleep(2)
    #   x,y,z,w = win32gui.GetWindowRect(1378600)
    #   pyautogui.click(x,y)


def printfile(filename,filenum):
    temp_Handle = win32gui.FindWindow('XLMAIN', None)
    win32gui.SetForegroundWindow(temp_Handle)
    #pyautogui.hotkey('ctrl', 'p')
    print("printing")


def unprotectfile(filename,filenum):
    pass


def verifyedit(filename,filenum):
    openfile(filename,filenum)
    if xlwings.Range('A3').value == 'Abc':
        reportlog("Data Matched",filenum)
    else:
        reportlog("Data not matched",filenum)
    closefile(filename,filenum)











