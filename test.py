# Import Library
from win32com import client
import os
a=os.getcwd()

# Opening Microsoft Excel
app = client.Dispatch("Excel.Application")
app.Interactive=False
app.Visible=False

# Read Excel File

workbook=app.Workbooks.Open(a+"\Resurs1")
workbook.ActiveSheet.ExportAsFixedFormat(0,a+"\Resurs1")
workbook.Close()
# Converting into PDF File