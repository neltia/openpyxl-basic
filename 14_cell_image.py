# 14_cell_image.py
# - Post 11. 셀에 이미지 삽입

from openpyxl import Workbook
from openpyxl.drawing.image import Image

wb = Workbook()
ws = wb.active
ws['A1'] = '산출물 예시'

# 이미지 객체 호출 및 속성 확인
img = Image('example_img.jpg')
print(dir(img))
print(img.anchor)
print(img.format)
print(img.path)
print(img.ref)

# 이미지 사이즈 설정
# img.height = 100
# img.width = 100

# 이미지 삽입
ws.add_image(img, 'A2')
# 이미지 셀 위치 확인
print(img.anchor)

# 파일 저장
wb.save('wb_image.xlsx')
