# pretreatment testing
import os
import xlwings as xw
import openpyxl
from openpyxl.utils import get_column_letter

wbxl=xw.Book('./out.xlsx')
for row in  range (1, 20):
    for col in range (10, 25):
        cell = get_column_letter(col) + str(row)
        wbxl.sheets['raw_copy'].range(cell).value = wbxl.sheets['raw_copy'].range(cell).value
wbxl.save()
os.system("C:\Windows\System32\\taskkill.exe /IM EXCEL.EXE /F")
