#!/usr/bin/env python 
# -*- coding:utf-8 -*-

from openpyxl import Workbook
from openpyxl.chart import (
    LineChart,
    Reference,
)

wb = Workbook()
ws = wb.active

rows = [
    ['ValueName', 'person1', 'person2', 'person3'],
    ['Pathfinding', 40, 30, 25],
    ['Aligning', 40, 25, 30],
    ['Empowering', 50, 30, 45],
    ['Modeling', 30, 25, 40],
]

for row in rows:
    ws.append(row)

data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=5)

# Chart with date axis
c2 = LineChart()
c2.style = 26  # 17 黑白

c2.add_data(data, titles_from_data=True)
labels = Reference(ws, min_col=1, min_row=2, max_row=5)
c2.set_categories(labels)

ws.add_chart(c2, "A61")

wb.save("../Resource/LineChart.xlsx")