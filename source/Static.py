#!/usr/bin/env python 
# -*- coding:utf-8 -*-

import os


rootDir = '../Resource/input'

ignoreFileName = [
    'RolesOfLeadershipDetail.xlsx',
    'RolesOfLeadershipOverview.xlsx',
]

source_col = 'C'
source_begin_row = 16
source_end_row = 79

template_detail_table_name = 'RolesOfLeadershipDetail.xlsx'
template_detail_table_path = os.path.join('../Resource/', template_detail_table_name)

template_overview_table_name = 'RolesOfLeadershipOverview.xlsx'
template_overview_table_path = os.path.join('../Resource/', template_overview_table_name)

targetRootDir = '../Resource/output'

overviewNameDict = {'Boss': {'Pathfindings': 'E12', 'Aligning': 'F12', 'Empowering': 'G12', 'Modeling': 'H12'},
                    'Peer': {'Pathfindings': 'E13', 'Aligning': 'F13', 'Empowering': 'G13', 'Modeling': 'H13'},
                    'DirectReport': {'Pathfindings': 'E14', 'Aligning': 'F14', 'Empowering': 'G14', 'Modeling': 'H14'},
                    'Self': {'Pathfindings': 'E11', 'Aligning': 'F11', 'Empowering': 'G11', 'Modeling': 'H11'}}

targetSheetArr = ['Pathfindings', 'Aligning', 'Empowering', 'Modeling']

targetEvaluateValueDict = {'Strongly Disagree': 0,
                           'Disagree': 20,
                           'Slightly Disagree': 40,
                           'Slightly Agree': 60,
                           'Agree': 80,
                           'Strongly Agree': 100}

"""
[col_position, can_direct_write_data]
['L', True]
"""
targetNameDict = {'DirectReport': ['L', True],
                  'Peer': ['M', False],
                  'Boss': ['N', True],
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


def clean_output_dir(dir_path):
    ls = os.listdir(dir_path)
    for i in ls:
        c_path = os.path.join(dir_path, i)
        if os.path.isdir(c_path):
            clean_output_dir(c_path)
        else:
            os.remove(c_path)




