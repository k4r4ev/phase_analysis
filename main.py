import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
import os

print("Phase analysis program. Sponsored by NATO & NASA")

workbook = openpyxl.load_workbook('data.xlsx')

sheet = workbook["test"]

# value = sheet['A1'].value
xvals = []
yvals = []
value = sheet.cell(1, 1).value
for row in range(2, 20):
    xvals.append(sheet.cell(row, 2).value)
    yvals.append(sheet.cell(row, 3).value)

print(xvals)
print(min(xvals))

chart = ScatterChart()
chart.title = "Квазицикл"
chart.style = 2
chart.x_axis.title = ''
chart.y_axis.title = ''
chart.legend = None

xvalues = Reference(sheet, min_col=2, min_row=2, max_row=20)
yvalues = Reference(sheet, min_col=3, min_row=2, max_row=20)

chart.x_axis.scaling.min = min(xvals) - 10
chart.y_axis.scaling.min = min(yvals) - 10
chart.x_axis.scaling.max = max(xvals) + 10
chart.y_axis.scaling.max = max(yvals) + 10

series = Series(yvalues, xvalues, title_from_data=True)
chart.layoutTarget = "inner"
chart.series.append(series)

sheet.add_chart(chart, "C15")

os.remove("data.xlsx")
workbook.save('data.xlsx')
