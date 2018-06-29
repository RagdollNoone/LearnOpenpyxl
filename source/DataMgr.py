#!/usr/bin/env python
# -*- coding:utf-8 -*-

import Static


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
            for index in self.peerDetailTable[name]:
                data = int(self.peerDetailTable[name][index] / 4)
                self.peerDetailTable[name][index] = data
                self.process_peer_score(name, "peer", data)
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

    def generate_average_score(self):
        for name in self.overViewTable:
            for identity in self.overViewTable[name]:
                data = int(self.overViewTable[name][identity] / 64)
                self.overViewTable[name][identity] = data
        return

    def generate_other_average_score(self):
        self.convertInput2Detail.generate_other_average(Static.targetRootDir)
        return

    def generate_overview_table(self):
        self.setData2Overview.generate_overview_table(Static.targetRootDir)