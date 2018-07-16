#!/usr/bin/env python 
# -*- coding:utf-8 -*-

import os
from BaseXlsx import BaseXlsx
from Base import MISSING
from Base import IDENTITY_DICT
from Base import get_data_mgr_instance


class InputXlsx(BaseXlsx):
    def __init__(self, file_path):
        BaseXlsx.__init__(self, file_path)

        self.identity = MISSING
        self.write_way = MISSING
        self.range = MISSING

        # self.active_self()

    def active_self(self, config_path):
        self.active_xlsx(config_path)

        self.set_write_way()
        self.range = self.get_sheet_range()
        self.identity = self.scan_identity()
        assert self.identity is not MISSING, \
            "identity is illegal file name should be one of \n" \
            "DirectReport Peer Boss Self + 2 + name.xlsx"

    def scan_identity(self):
        base_name = os.path.basename(self.file_path)
        pre = base_name.split("2")[0]

        identity = MISSING
        for key in IDENTITY_DICT:
            check_key = key.replace(" ", "")  # remove blank in Direct Report
            if check_key in pre:
                identity = key
                return identity

        return MISSING

    def set_write_way(self):
        assert self.identity is not MISSING, \
            "current identity is illegal\n" \
            "may be you need scan_identity first"

        self.write_way = IDENTITY_DICT[self.identity]
        return

    def scan_table(self):
        if self.write_way:
            begin_cell = self.get_cell_by_content(self.identity, need_cache=True)

            assert begin_cell is not MISSING

            for offset in range(self.range):
                need_process_cell = self.get_cell_by_cell_and_offset(begin_cell, row_offset=offset)
                data_mgr = get_data_mgr_instance()
                data_mgr.inputxlsx_data_2_detailxlsx(self.file_path,
                                                     self.identity,
                                                     offset,
                                                     need_process_cell)
        else:

            pass


ip = "../resource/input/a/Boss2a.xlsx"
cp = "../resource/InputExcelConfig.txt"

obj = InputXlsx(ip)
obj.active_self(cp)
obj.scan_table()



