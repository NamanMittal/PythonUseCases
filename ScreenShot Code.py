# -*- coding: utf-8 -*-
"""
Created on Thu Sep  5 15:02:19 2019

@author: TA0056
"""
#import win32gui
import win32com.client
import time
import datetime 

Application = win32com.client.DispatchEx("Excel.Application")
Application.Visible = 1
wb = Application.Workbooks.Open("C:\\Users\\ta0056\\Desktop\\Pivot.xlsx")
wb.RefreshAll()
Application.CalculateUntilAsyncQueriesDone()

time.sleep(5)

wb.Save()
wb.Close(SaveChanges=1)
Application.Quit()

import subprocess
subprocess.Popen([r'C:\Users\ta0056\Desktop\test.bat'])
time.sleep(5)

import pyscreenshot as ImageGrab
im = ImageGrab.grab(bbox=(0,0,1365,730)) # X1,Y1,X2,Y2
im.save('screenshot '+datetime.datetime.now().strftime("%d%m%Y_%H%M%S")+'.png')
#im.show()
p = subprocess.Popen([r'C:\Users\ta0056\Desktop\test - close.bat'])





