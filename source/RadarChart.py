#!/usr/bin/env python 
# -*- coding:utf-8 -*-

from openpyxl import Workbook
from openpyxl.chart import (
    RadarChart,
    Reference,
    Series,
)

wb = Workbook()
ws = wb.active

rows = [
    ['Relationships', 'Pathfinding', 'Aligning', 'Modeling', 'Empowering'],
    ['Self', 97, 95, 98, 90],
    ['Boss', 61, 65, 66, 66],
    ['Peer', 65, 62, 65, 61],
    ['Direct Report', 100, 95, 89, 88],
]

for row in rows:
    ws.append(row)

chart = RadarChart()
chart.type = "marker"

labels = Reference(ws, min_col=2, min_row=1, max_col=5)

for i in range(2, 6):
    values = Reference(ws, min_row=i, min_col=1, max_col=5)
    series = Series(values, title_from_data=True)
    chart.series.append(series)

chart.set_categories(Reference(ws, min_col=2, min_row=1, max_col=5))

chart.style = 26
chart.y_axis.delete = True

ws.add_chart(chart, "A17")

wb.save("../Resource/RadarChat.xlsx")