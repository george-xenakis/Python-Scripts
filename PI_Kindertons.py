from datetime import datetime
from os.path import getmtime
import shutil
import os
import win32com.client

xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = True
wb = xl.Workbooks.Open(r'\\cru-file-01\Admin\PI_Import_Kindertons\PI_Kindertons_Template.xlsm')
xl.CalculateUntilAsyncQueriesDone()
wb.Application.Run("PI_Kindertons_Template.xlsm!Module1.Automacro")
xl.CalculateUntilAsyncQueriesDone()
xl.DisplayAlerts = False
wb.Close()
xl.DisplayAlerts = False
xl.Quit()