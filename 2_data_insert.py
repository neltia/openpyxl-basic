# 2_data_insert.py
from openpyxl import load_workbook

wb = load_workbook('wb.xlsx')
ws = wb['Sheet']

print(ws['A1'].value)
ws['A1'] = "TEST DATA"
print(ws['A1'].value)

print(ws.cell(row=1, column=1))

wb.save('wb_insert.xlsx')

# print(type(ws.cell(row=1, column=1)))
# print(dir(ws.cell(row=1, column=1)))
