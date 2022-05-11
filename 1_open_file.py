# 1_open_file.py
from openpyxl import load_workbook

wb = load_workbook('wb.xlsx')
print(wb.sheetnames)
