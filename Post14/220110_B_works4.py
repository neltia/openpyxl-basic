# 필요 라이브러리 호출
import openpyxl
from openpyxl import load_workbook
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.chart.layout import Layout, ManualLayout

# 제1작업 시트 불러오기
wb = load_workbook('result-2201010B-openpyxl_part1.xlsx')
ws = wb["제1작업"]

# 제4작업 차트 시트 생성
ws4 = wb.create_chartsheet("제4작업")

# refer
refer_names = "제1작업!$C$5:$C$12"
refer_money = "제1작업!$H$5:$H$12"
refer_time =  "제1작업!$F$5:$F$12"

# bar chart
chart_bar = BarChart()
refer_name_string = Reference(ws4, range_string=refer_names)
refer_name_position = Reference(ws4, min_col=3, min_row=5, max_row=12)
refer_money_string = Reference(ws4, range_string=refer_money)

v1 = Series(refer_name_string)
data1 = Series(refer_money_string, title="급여(단위: 원)")

chart_bar.series = (v1, data1)
chart_bar.set_categories(refer_name_string)
chart_bar.title = '배송부 및 식료사업부 급여 현황'

# line cart
chart_line = LineChart()
refer3 = Reference(ws4, range_string=refer_time)

data2 = Series(refer3, title="근속기간")
chart_line.series = (v1, data2)
chart_line.y_axis.axId = 200
chart_line.y_axis.number_format = '#,##0"년"'

# chart add
chart_bar.y_axis.majorGridlines = None
chart_bar.y_axis.crosses = "max"
chart_bar += chart_line

# legend positoin
'''
chart_line.legend.position = 'tr'
chart_bar.legend.position = "tr"
chart_bar.legend.layout = Layout(
    manualLayout=ManualLayout(
        yMode='edge',
        xMode='edge',
        x=0, y=0.9,
        h=0.1, w=0.5
    )
)
'''

ws4.add_chart(chart_bar)

# 완료 데이터 저장
wb.save("result-2201010B-openpyxl_part4.xlsx")
