from openpyxl import workbook, load_workbook
import win32com
import os
a=os.getcwd()
a=a+"\wyniki"
wb=load_workbook(filename="Resurswzor.xlsx")
ws=wb.active
ws["l8"]="cos"
print(a)
f=r"C:\Users\Dell\Desktop"
wb.save("j.xlsx",f)
