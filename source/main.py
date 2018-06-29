#!/usr/bin/env python
# -*- coding:utf-8 -*-
import ConvertInput2Detail
import DataMgr
import Static

# 清空output的数据

# init
dm = DataMgr.DataMgr()
ci2d = ConvertInput2Detail.ConvertInput2Detail()

# comtain
dm.set_convert_input_2_detail(ci2d)
ci2d.set_data_mgr(dm)

# 执行ConvertInput2Detail
ci2d.convert(Static.rootDir)

# 执行DataMgr分析Detail数据

# 生成Overview文件

# 生成图像

# excel转pdf输出