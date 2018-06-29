#!/usr/bin/env python
# -*- coding:utf-8 -*-

import os
import shutil


class ConvertInput2Detail():
    rootDir = '../Resource/input'

    ignoreFileName = [
        'RolesOfLeadershipDetail.xlsx',
        'RolesOfLeadershipOverview.xlsx',
    ]

    targetEvaluateValueDict = {'Strongly Disagree': 0,
                               'Disagree': 20,
                               'Slightly Disagree': 40,
                               'Slightly Agree': 60,
                               'Agree': 80,
                               'Strongly Agree': 100}

    source_col = 'C'
    source_begin_row = 16
    source_end_row = 79

    targetRootDir = '../Resource/output'

    """
    [col_position, can_direct_write_data]
    ['L', True]
    """
    targetNameDict = {'DirectReport': ['L', False],
                      'Peer': ['M', False],
                      'Boss': ['N', False],
                      'OthersAverage': ['O', True],
                      'Self': ['P', True]}

    """
    [begin_row_index, end_row_index, question_index_begin, question_index_end]
    [2, 16, 0, 14, 0]
    """
    targetSheetDict = {'Pathfindings': [2, 16, 0, 14],
                       'Aligning': [2, 20, 15, 33],
                       'Empowering': [2, 14, 34, 46],
                       'Modeling': [2, 18, 47, 63]}

    template_detail_table_name = 'RolesOfLeadershipDetail.xlsx'
    template_detail_table_path = os.path.join('../Resource/', template_detail_table_name)

    def __init__(self):
        return

    def convert(self, file_dir):
        file_list = os.listdir(file_dir)

        for i in range(0, len(file_list)):
            path = os.path.join(file_dir, file_list[i])
            base_name = os.path.basename(path)

            if self.check_is_ignore_file(base_name):
                continue

            if os.path.isdir(path):
                self.convert(path)

            file_format_arr = os.path.splitext(path)
            if file_format_arr[1] == ".xlsx":

                identity = self.get_file_identity(base_name)
                if identity == "illegal identity":
                    print("Can not recognize the file " + base_name + " identity")
                    return

                can_direct_write = self.check_can_direct_write(identity)
                target_file_path = self.prepare_for_output(path)
                self.data_operate(path, identity, can_direct_write, target_file_path)
        return

    def check_evaluate_string_legal(self, evaluate):
        result = False

        for key in self.targetEvaluateValueDict:
            if key == evaluate:
                return True

        return result

    def get_evaluate_value(self, evaluate):
        for key in self.targetEvaluateValueDict:
            if key == evaluate:
                return self.targetEvaluateValueDict[key]

        return -1

    def check_is_ignore_file(self, name):
        for i in range(len(self.ignoreFileName)):
            if name == self.ignoreFileName[i]:
                return True

        return False

    def check_can_direct_write(self, name):
        if not (name in self.targetNameDict):
            return -1

        return self.targetNameDict[name][1]

    @staticmethod
    def get_file_identity(file_name):
        pre = file_name.split('2')[0]
        if "Boss" in pre:
            return "Boss"

        if "Peer" in pre:
            return "Peer"

        if "DirectReport" in pre:
            return "DirectReport"

        if "Self" in pre:
            return "Self"

        return "illegal identity"

    """
    生成输出目录和文件
    """
    def prepare_for_output(self, path):
        parent_folder_name = os.path.dirname(path)
        target_folder_path = parent_folder_name.replace("input", "output")
        target_file_path = os.path.join(target_folder_path, self.template_detail_table_name)

        if not os.path.exists(target_folder_path):
            os.makedirs(target_folder_path)

        if not os.path.exists(target_file_path):
            shutil.copyfile(self.template_detail_table_path, target_file_path)

        return target_file_path

    def data_operate(self, path, evaluate_identity, can_direct_write, target_file_path):
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb.active

        for j in range(0, self.source_end_row - self.source_begin_row + 1):
            read_unit = self.source_col + str(j + self.source_begin_row)

            if self.check_evaluate_string_legal(ws[read_unit].value):
                # print(ws[read_unit].value)
                if can_direct_write:
                    self.write_data(evaluate_identity, target_file_path, ws[read_unit].value, j)
                # else:
                #     # TODO
                #     print("Call DataMgr function")

            else:
                print("illegal evaluate in file " + os.path.abspath(path) + " unit " + read_unit)
                return

        return

    def write_data(self, evaluate_identity, target_file_path, value, index):
        from openpyxl import load_workbook
        wb_write = load_workbook(target_file_path)
        write_sheet, write_unit = self.get_target_sheet_and_unit(evaluate_identity, index)
        ws_write = wb_write[write_sheet]

        ws_write[write_unit] = self.get_evaluate_value(value)
        wb_write.template = False
        wb_write.save(target_file_path)

    def get_target_sheet_and_unit(self, evaluate_identity, index):
        for key in self.targetSheetDict:
            if index >= self.targetSheetDict[key][2] and index <= self.targetSheetDict[key][3]:
                offset = index - self.targetSheetDict[key][2]
                target_unit = str(self.targetNameDict[evaluate_identity][0]) + str(self.targetSheetDict[key][0] + offset)
                return key, target_unit

        return "illegal_sheet", -1
