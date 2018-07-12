#!/usr/bin/env python 
# -*- coding:utf-8 -*-

from Base import MISSING
from openpyxl import load_workbook

class BaseXlsx(object):

    def __init__(self, file_path):
        self.file_path = file_path
        self.wb = MISSING
        self.current_ws = MISSING
        self.row_name = {}
        self.col_name = {}
        self.top_left_str = get_top_left_str()

    def load_sheet_by_name(self, sheet_name):
        result = MISSING
        if self.wb is MISSING:
            self.wb = load_workbook(self.flie_path)

        result = self.wb[sheet_name]
        self.current_ws = result

        return result

    def write_value(self, unit_name, value):
        assert self.wb is not MISSING
        assert self.current_ws is not MISSING

        self.current_ws[unit_name] = value
        self.wb.template = False
        self.wb.save(self.flie_path)

    def get_col_index_by_name(self, col_name):
        pass

    def get_row_index_by_name(self, row_name):
        pass

    @staticmethod
    def get_unit(row_index, col_index):
        return str(row_index) + str(col_index)



