#!/usr/bin/env python 
# -*- coding:utf-8 -*-

import os
import Static


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

            if base_name == Static.template_overview_table_name:
                from openpyxl import load_workbook
                to_wb = load_workbook(path)
                from_wb = load_workbook(Static.template_detail_table_name)

                for i in range(len(Static.targetSheetArr)):
                    from_ws = from_wb[Static.targetSheetArr[i]]

                    # for i in range(Static.targetSheetDict[key][0], Static.targetSheetDict[key][1]):
                    #     direct_value = ws[('L' + str(i))]
                    #     peer_value = ws[('M' + str(i))]
                    #     boss_value = ws[('N' + str(i))]
                    #     value = (direct_value + peer_value + boss_value) / 3
                    #     ws[('O' + str(i))] = value
                    #     wb.template = False
                    #     wb.save(path)

        return