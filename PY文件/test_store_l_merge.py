#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试：用5G工具，只跑_store_l的合并逻辑，输出段分布
"""

import sys, os, json, re
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any
import numpy as np
import pandas as pd

# 直接导入工具中的核心类
sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')
from importlib import util as import_util
spec = import_util.spec_from_file_location("tool",
    "/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py")
tool = import_util.module_from_spec(spec)
spec.loader.exec_module(tool)

# 读取联通 UCM
uproc = tool.UCMProcessor({'dl_peak_limit': 900, 'ul_peak_limit': 160})
paths = [
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115328-电信.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115334-电信.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336-联通.xlsx'
]

params = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10, 'match_mode': 'auto'}
agg = uproc.parse_mmf(paths, params)
uproc._fn = 'mmf_test'

# 运行自动匹配
round_results = uproc._match_auto(agg, params, 50, 10)

print("\n============ 结果摘要 ============")
from collections import Counter
key_counts = Counter()
total_rounds = 0
for rd in round_results:
    total_rounds += 1
    for k in rd:
        key_counts[k] += 1

print(f"总轮次: {total_rounds}")
print(f"各业务次数: {dict(key_counts)}")

# 详细：store_l 信息
print("\n============ store_l 详情 ============")
for i, rd in enumerate(round_results):
    for k, s in rd.items():
        if k == 'store_l':
            print(f"  轮{i}: flow={s['flow_mb']:.0f}MB dur={s['duration']:.1f}s rate={s['rate']:.1f}Mbps "
                  f"start={s['start_t']} end={s['end_t']}")

print("\n============ store_s 详情 ============")
for i, rd in enumerate(round_results):
    for k, s in rd.items():
        if k == 'store_s':
            print(f"  轮{i}: flow={s['flow_mb']:.0f}MB dur={s['duration']:.1f}s rate={s['rate']:.1f}Mbps "
                  f"start={s['start_t']}")

# 所有段（debug）
print("\n============ 所有段 flow/分类 ============")
for i, rd in enumerate(round_results):
    for k, s in sorted(rd.items()):
        print(f"  轮{i} {k:8s}: flow={s['flow_mb']:6.0f}MB dur={s['duration']:6.1f}s rate={s['rate']:6.1f}Mbps "
              f"start={s['start_t'].strftime('%H:%M:%S')}")