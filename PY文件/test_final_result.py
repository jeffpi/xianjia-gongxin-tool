#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证 gap<10合并 + 新分类 的最终结果（完整跑match → 输出结果）
对比修改前后的差异
"""

import numpy as np
import pandas as pd
from importlib import util as import_util

spec = import_util.spec_from_file_location('tool', '5G用户级公共监控速率统计工具-V1.04.py')
tool = import_util.module_from_spec(spec)
spec.loader.exec_module(tool)

uproc = tool.UCMProcessor({'dl_peak_limit': 900, 'ul_peak_limit': 160})
paths = ['联通/mmf20260703115328-电信.xlsx','联通/mmf20260703115334-电信.xlsx','联通/mmf20260703115336-联通.xlsx']
params = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10, 'match_mode': 'auto'}
agg = uproc.parse_mmf(paths, params)
uproc._fn = 'mmf_test'
uproc._fn = 'test'

# 完整跑 match → 呼叫详情 + 统计
call_df, stats = uproc.match(tool.DEFAULT_PLAN, params)

# 从stats获取
store_s_count = stats.get('应用商店_小包', {}).get('测试次数', '?')
store_l_count = stats.get('应用商店_大包', {}).get('测试次数', '?')
ftp_dl_count = stats.get('FTP下载', {}).get('测试次数', '?')
ftp_ul_count = stats.get('FTP上传', {}).get('测试次数', '?')

print(f"统计汇总: store_s 测试次数={store_s_count}")
print(f"统计汇总: store_l 测试次数={store_l_count}")
print(f"统计汇总: ftp_dl 测试次数={ftp_dl_count}")
print(f"统计汇总: ftp_ul 测试次数={ftp_ul_count}")

# 统计各业务次数（从round_results）
round_results = uproc._last_round_results
from collections import Counter
key_counts = Counter()
for rd in round_results or []:
    for k in rd:
        key_counts[k] += 1
print(f"\n自动识别结果:")
print(f"  总轮次: {len(round_results or [])}")
print(f"  各业务次数: {dict(key_counts)}")
print(f"  store_l 期望~26, 实际={key_counts.get('store_l', 0)}")
print(f"  store_s 期望~26, 实际={key_counts.get('store_s', 0)}")
print(f"  ftp_dl  期望~26, 实际={key_counts.get('ftp_dl', 0)}")
print(f"  ftp_ul  期望~26, 实际={key_counts.get('ftp_ul', 0)}")