# 10_ws_rowscols.py
# - Post 8. 행열 삽입과 삭제

from openpyxl import load_workbook
from openpyxl.styles import Alignment

# 5_functions.py 파일의 결과 파일 경로 입력 후 파일 load
path = "./preview_resultfile"
wb = load_workbook(f'{path}/wb_multiple_func.xlsx')
ws = wb['Sheet']

# 행열 삽입
ws.insert_rows(5, 1) # 5행에 1개의 행 삽입
ws.insert_cols(3, 2) # 3열에 2개의 열 삽입

# 중간 확인
# wb.save("test.xlsx")

# 행열 삭제
ws.delete_rows(3, 2) # 3열에 2개의 행 삭제
ws.delete_cols(5, 1) # 5행에 1개의 열 삭제
wb.save("test.xlsx")
