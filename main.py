import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference


#read data
df = pd.read_excell('data.xlsx')

print(df.head())


filtered = df[df['Garage'] == 'Engen Killner Park']

quartely_sales = pd.pivot_table(filterd, index = filtered['date'].dt.quarter, columns='Type', values = 'TotalCost', aggfunc='sum')

