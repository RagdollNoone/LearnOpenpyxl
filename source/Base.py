#!/usr/bin/env python 
# -*- coding:utf-8 -*-

from DataMgr import DataMgr

MISSING = object()

SPLIT_CHAR_EQUAL = "="
SPLIT_CHAR_COMMA = ","

IDENTITY_DICT = {"Direct Report": True,
                 "Peer": False,
                 "Boss": True,
                 "Self": True}

DETAIL_FILE_NAME = "RolesOfLeadershipDetail.xlsx"
OVERVIEW_FILE_NAME = "RolesOfLeadershipOverview.xlsx"

DETAIL_CONFIG_FILE = "../resource/DetailExcelConfig.txt"
OVERVIEW_CONFIG_FILE = "../resource/OverviewExcelConfig.txt"

data_mgr = DataMgr()


def get_data_mgr_instance():
    return data_mgr

# copy

# makedir








