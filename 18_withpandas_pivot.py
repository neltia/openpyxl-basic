# 18_withpandas_pivot.py
# - Post 13. 판다스(pandas)와 같이 사용하기

# 라이브러리 호출
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime
import time

df = pd.read_excel("example_data.xlsx", sheet_name="Sheet1")
df.columns = df.iloc[0, :]
df = df.iloc[1:, :]

df['거래일자'] = pd.to_datetime(df["거래일자"], format='%Y-%m-%d')
df['월별'] = df['거래일자'].dt.strftime('%m')

df.pivot = df.pivot_table(index="거래 품명",
    values="매출액",
    columns="월별",
    aggfunc='sum'
)

book = load_workbook('example_data.xlsx')
writer = pd.ExcelWriter('wb_pivot.xlsx', engine='openpyxl') # index이거 pandas에 없는데.
writer.book = book
time.sleep(3)
df.pivot.to_excel(writer, sheet_name="Pivot")
writer.save()

