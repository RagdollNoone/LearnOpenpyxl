#!/usr/bin/env python
# -*- coding:utf-8 -*-
import ConvertInput2Detail
import DataMgr
import SetData2Overview

from Static import targetRootDir
from Static import clean_output_dir
from Static import rootDir

scan_data = False
generate_chart = False

# init
dm = DataMgr.DataMgr()
ci2d = ConvertInput2Detail.ConvertInput2Detail()
sd2o = SetData2Overview.SetData2Overview()

# 清空output的数据
if scan_data:
    clean_output_dir(targetRootDir)
    print("##########################")
    print("clean_output_dir is finish")
    print("##########################")

# contain
dm.set_convert_input_2_detail(ci2d)
dm.set_set_data_2_overview(sd2o)
ci2d.set_data_mgr(dm)

# 执行ConvertInput2Detail
if scan_data:
    ci2d.convert(rootDir)
    print("##########################")
    print("convert is finish")
    print("##########################")

# 执行DataMgr分析Detail数据
if scan_data:
    dm.generate_peer_average_score()
    print("##########################")
    print("generate_peer_average_score is finish")
    print("##########################")

if scan_data:
    dm.generate_identity_average_score()
    print("##########################")
    print("generate_identity_average_score is finish")
    print("##########################")

if scan_data:
    dm.generate_other_average_score()
    print("##########################")
    print("generate_other_average_score is finish")
    print("##########################")

# 生成Overview文件
dm.generate_overview_table()
print("##########################")
print("generate_overview_table is finish")
print("##########################")

# 生成图像
if generate_chart:
    sd2o.generate_line_chart(targetRootDir)
    sd2o.generate_radar_chart(targetRootDir)

print("##########################")
print("Successfully!~")
print("##########################")

# excel转pdf输出
