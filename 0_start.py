# 0_start.py
from openpyxl import Workbook
wb = Workbook()

ws = wb.active # 기본 시트 가져오기
ws1 = wb.create_sheet("Sheet1") # 가장 뒤에 시트 생성 (기본값)

wb.save('wb.xlsx')
