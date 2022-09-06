# 필요 라이브러리 호출
# - load workbook
from openpyxl import load_workbook
# - chart generate
from openpyxl.chart import LineChart, BarChart, Reference, Series
# - chart font
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font
# - chart
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
# - chart labels
from openpyxl.chart.label import DataLabelList
# - chart background color setting
from openpyxl.chart.shapes import GraphicalProperties
# -
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
refer_money_position = Reference(ws4,
    min_col=6,
    min_row=5,
    max_row=12
)

v1 = Series(refer_name_string)
data_money_string = Series(refer_money_string, title="급여(단위: 원)")
data_money_position = (Series(refer_money_position),)

chart_bar.series = (data_money_string,)
chart_bar.set_categories(refer_name_string)
chart_bar.y_axis.number_format = '#,##'

# data labels
chart_bar.dataLabels = DataLabelList()
chart_bar.dataLabels.showVal = True
print(dir(chart_bar))

# line cart
chart_line = LineChart()
refer3 = Reference(ws4, range_string=refer_time)

data2 = Series(refer3, title="근속기간")
chart_line.series = (data2,)
chart_line.set_categories(refer_name_string)
chart_line.y_axis.axId = 200
chart_line.y_axis.number_format = '#,##0"년"'
chart_line.title = '배송부 및 식료사업부 급여 현황'

# Style chart: X and Y axes numbers
def set_chart_title_size(chart, size=1100):
    cp = CharacterProperties(latin=font, sz=size)
    paraprops = ParagraphProperties(defRPr=cp)
    # paraprops.defRPr = CharacterProperties(latin=font, sz=size)

    for para in chart.title.tx.rich.paragraphs:
        para.pPr=paraprops

font = Font(typeface='굴림')
# - 11 point size
size = 1100
# - Not bold
cp = CharacterProperties(latin=font, sz=size, b=False)
pp = ParagraphProperties(defRPr=cp)
rtp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
# - y축, txPr=textProperties
chart_bar.y_axis.txPr = rtp
chart_line.y_axis.txPr = rtp
# - x 축
chart_line.x_axis.txPr = rtp
# - 범례
chart_line.legend.txPr = rtp
# - 제목
set_chart_title_size(chart_line, size=2000)
# - 데이터 라벨
chart_bar.dataLabels.txPr = rtp

# Marker line chart
s1 = chart_line.series[0]
s1.marker.symbol = "diamond"

# chart add
chart_bar.y_axis.majorGridlines = None
chart_bar.style = 10
chart_bar.y_axis.crosses = "max"
chart_line += chart_bar

# chart background color
props = GraphicalProperties(solidFill="f2f2f2")
# print(help(GraphicalProperties()))
chart_line.plot_area.graphicalProperties = props

# legend positoin
# - https://openpyxl.readthedocs.io/en/latest/charts/chart_layout.html
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

ws4.add_chart(chart_line)

# 완료 데이터 저장
wb.save("result-2201010B-openpyxl_part4.xlsx")
