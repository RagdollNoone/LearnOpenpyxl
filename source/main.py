#!/usr/bin/env python
# -*- coding:utf-8 -*-
import ConvertInput2Detail
import DataMgr
import SetData2Overview

from Static import targetRootDir
from Static import clean_output_dir
from Static import rootDir

# init
dm = DataMgr.DataMgr()
ci2d = ConvertInput2Detail.ConvertInput2Detail()
sd2o = SetData2Overview.SetData2Overview()

# 清空output的数据
clean_output_dir(targetRootDir)

# contain
dm.set_convert_input_2_detail(ci2d)
dm.set_set_data_2_overview(sd2o)
ci2d.set_data_mgr(dm)

# 执行ConvertInput2Detail
ci2d.convert(rootDir)

# 执行DataMgr分析Detail数据
dm.generate_peer_average_score()
dm.generate_average_score()
dm.generate_other_average_score()

# 生成Overview文件
dm.generate_overview_table()

# 生成图像

# excel转pdf输出