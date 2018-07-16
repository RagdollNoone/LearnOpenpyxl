#!/usr/bin/env python
# -*- coding:utf-8 -*-

import os

from Static import targetRootDir
from Static import template_detail_table_name
from Base import DETAIL_FILE_NAME
from Base import OVERVIEW_FILE_NAME
from Base import DETAIL_CONFIG_FILE
from Base import OVERVIEW_CONFIG_FILE


class DataMgr():
    convertInput2Detail = None
    setData2Overview = None
    peerDetailTable = {}
    overViewTable = {}

    def __init__(self):
        return

    def set_convert_input_2_detail(self, obj):
        self.convertInput2Detail = obj
        return

    def set_set_data_2_overview(self, obj):
        self.setData2Overview = obj

    def process_peer_score(self, name, index, value):
        if not (name in self.peerDetailTable):
            self.peerDetailTable[name] = {}
            self.peerDetailTable[name][index] = value
        else:
            if not (index in self.peerDetailTable[name]):
                self.peerDetailTable[name][index] = value
            else:
                data = self.peerDetailTable[name][index]
                self.peerDetailTable[name][index] = data + value
        return

    def generate_peer_average_score(self):
        for name in self.peerDetailTable:
            dir_path = os.path.join(targetRootDir, name)
            file_path = os.path.join(dir_path, template_detail_table_name)
            print("generate_peer_average_score file name is " + file_path)

            for index in self.peerDetailTable[name]:
                data = round(self.peerDetailTable[name][index] / 4)
                self.peerDetailTable[name][index] = data

                self.process_score(name, "Peer", data)
                self.convertInput2Detail.write_data("Peer", file_path, data, index)
        return

    def process_score(self, name, identity, value):
        if not (name in self.overViewTable):
            self.overViewTable[name] = {}
            self.overViewTable[name][identity] = value
        else:
            if not (identity in self.overViewTable[name]):
                self.overViewTable[name][identity] = value
            else:
                data = self.overViewTable[name][identity]
                self.overViewTable[name][identity] = data + value
        return

    def generate_identity_average_score(self):
        for name in self.overViewTable:
            for identity in self.overViewTable[name]:
                data = round(self.overViewTable[name][identity] / 64)
                self.overViewTable[name][identity] = data
        return

    def generate_other_average_score(self):
        self.convertInput2Detail.generate_other_average(targetRootDir)
        return

    def generate_overview_table(self):
        self.setData2Overview.generate_overview_table(targetRootDir)

    #############################################
    detailxlsx_dict = {}

    # static
    def inputxlsx_data_2_detailxlsx(self, file_path, identity, question_id, cell):
        detail_obj = self.get_detailxlsx_through_inputxlsx_path(file_path)
        target_cell = detail_obj.get_cell_by_question_number(question_id, identity)
        detail_obj.write_value(target_cell, cell.value)

    def get_detailxlsx_through_inputxlsx_path(self, file_path):
        import os

        base_name = os.path.basename(file_path)
        be_evaluate_name = base_name.split("2")[1]
        if be_evaluate_name in self.detailxlsx_dict:
            return self.detailxlsx_dict[be_evaluate_name]

        dir_name = os.path.dirname(file_path)
        target_dir_name = dir_name.replace("input", "output")
        detail_path = os.path.join(target_dir_name, DETAIL_FILE_NAME)

        from DetailXlsx import DetailXlsx
        obj = DetailXlsx(detail_path)
        obj.active_self(DETAIL_CONFIG_FILE)
        self.detailxlsx_dict[be_evaluate_name] = obj

        return obj
