from win32com import client
import os
sciezka=os.getcwd()
sciezkawynik=os.getcwd()
sciezkawynik+="/wyniki"
excel = client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

sheets = excel.Workbooks.Open(sciezka+"\Resurs1")
work_sheets = sheets.Worksheets[0]

work_sheets.ExportAsFixedFormat(0, f'{sciezkawynik}\Resurs1')
excel.quit()
