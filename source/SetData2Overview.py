#!/usr/bin/env python 
# -*- coding:utf-8 -*-

import os

from Static import template_overview_table_name
from Static import template_detail_table_name
from Static import targetSheetArr
from Static import targetSheetDict
from Static import targetNameDict


class SetData2Overview():
    def __init(self):
        return

    def generate_overview_table(self, file_dir):
        file_list = os.listdir(file_dir)

        for i in range(0, len(file_list)):
            path = os.path.join(file_dir, file_list[i])
            base_name = os.path.basename(path)

            if os.path.isdir(path):
                self.generate_overview_table(path)

            if base_name == template_overview_table_name:
                from openpyxl import load_workbook
                to_wb = load_workbook(path)
                from_wb = load_workbook(template_detail_table_name)

                for j in range(len(targetSheetArr)):
                    sheet_name = targetSheetArr[j]
                    from_ws = from_wb[sheet_name]
                    begin_row = targetSheetDict[sheet_name][0]
                    end_row = targetSheetDict[sheet_name][1]

                    peer_total = 0
                    boss_total = 0
                    self_total = 0
                    direct_report_total = 0

                    for k in range(begin_row, end_row):
                        peer_total = peer_total + from_ws[self.get_unit('Peer', k)]
                        boss_total = boss_total + from_ws[self.get_unit('Boss', k)]
                        self_total = self_total + from_ws[self.get_unit('Self', k)]
                        direct_report_total = direct_report_total + from_ws[self.get_unit('DirectReport', k)]

                    number = end_row - begin_row + 1
                    peer_average = int(peer_total / number)
                    boss_average = int(boss_total / number)
                    self_average = int(self_total / number)
                    direct_report_average = int(direct_report_total / number)

                    to_ws = to_wb.active
                    to_ws[targetNameDict['Peer'][sheet_name]] = peer_average
                    to_ws[targetNameDict['Boss'][sheet_name]] = boss_average
                    to_ws[targetNameDict['Self'][sheet_name]] = self_average
                    to_ws[targetNameDict['DirectReport'][sheet_name]] = direct_report_average

                    to_wb.template = False
                    to_wb.save(path)
        return

    @staticmethod
    def get_unit(identity, index):
        col = targetNameDict[identity]
        result = col + str(index)
        return result
