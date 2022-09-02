# 필요 라이브러리 호출
from turtle import title
import openpyxl
from openpyxl import load_workbook
from openpyxl.chart import LineChart, BarChart, Reference, Series

# 제1작업 시트 불러오기
wb = load_workbook('result-2201010B-openpyxl_part1.xlsx')
ws = wb["제1작업"]

ws4 = wb.create_chartsheet("제4작업")

chart_bar = BarChart()

refer_names = "제1작업!$C$5:$C12"
refer_money = "제1작업!$H$5:$H12"

refer1 = Reference(ws4, range_string=refer_names)
refer2 = Reference(ws4, range_string=refer_money)
data1 = Series(refer1, title="이름")
data2 = Series(refer2, title="급여(단위: 원)")

print(dir(data1))

chart_bar.series = (data1, data2)
chart_bar.set_categories(refer1)

ws4.add_chart(chart_bar)

# 완료 데이터 저장
wb.save("result-2201010B-openpyxl_part4.xlsx")
