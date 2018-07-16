#!/usr/bin/env python 
# -*- coding:utf-8 -*-

from BaseXlsx import BaseXlsx
from Base import MISSING


class DetailXlsx(BaseXlsx):
    def __init__(self, file_path):
        BaseXlsx.__init__(self, file_path)

        self.pre_find_cell = MISSING

        pass

    def active_self(self, config_path):
        self.active_xlsx(config_path)

        pass

    def get_cell_by_question_number(self, question_id, identity):
        """
        :param question_id: question index number
        :param identity: one of boss, director report, peer, self
        :return: openpyxl struct "cell"
        """

        top_left_cell = self.get_cell_by_content(self.top_left_str)
        tc = top_left_cell.col_idx
        pr = self.pre_find_cell.row

        identity_cell = self.get_cell_by_content(identity)

        if self.current_ws.cell(row=pr, col=tc).value == question_id:
            result = self.current_ws.cell(row=pr, col=identity_cell.col_idx)
            return result

        if (pr + 1) > self.current_ws.max_row:
            self.get_next_sheet()
            return self.get_cell_by_question_number(question_id, identity)

        if self.current_ws.cell(row=(pr + 1), col=tc).value == question_id:
            result = self.current_ws.cell(row=pr, col=identity_cell.col_idx)
            return result

        result_row_cell = MISSING
        length = len(self.sheet_name_arr)
        for find_sheet_num in range(length):
            if find_sheet_num == length:
                break

            result_row_cell = self.get_cell_by_content(question_id)
            if result_row_cell is MISSING:
                self.get_next_sheet()

        if result_row_cell is MISSING:
            return MISSING

        result = self.current_ws.cell(row=result_row_cell.row, col=tc)

        return result

