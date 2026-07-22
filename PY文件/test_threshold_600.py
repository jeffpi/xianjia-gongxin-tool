#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证方案：将store_l分类阈值从1500降到600，看结果
"""

import sys, os, json, re
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any
import numpy as np
import pandas as pd

sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')
from importlib import util as import_util
spec = import_util.spec_from_file_location("tool",
    "/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py")
tool = import_util.module_from_spec(spec)
spec.loader.exec_module(tool)

# 备份原始_classify_seg方法
orig_classify = tool.UCMProcessor._classify_seg

def new_classify(self, s, segs, idx):
    d = s['direction']; flow = s['flow_mb']; dur = s['duration']; rate = s['rate']
    if d == '下行':
        if flow <= 50:
            return None
        if 50 < flow < 350:
            return 'store_s'
        if 350 <= flow < 700:
            return 'ftp_dl'
        if 700 < flow < 600:  # 永远不会走到这里，因为700<flow<600不可能
            return 'ftp_dl'
        # 关键改动：flow>=600就判store_l，而不是flow>=1500
        if flow >= 600:
            return 'store_l'
        return 'ftp_dl'
    else:  # 上行
        if flow < 20:
            return 'wx_s'
        if 20 <= flow <= 80:
            return 'ftp_ul'
        if 80 < flow < 200:
            return 'ftp_ul' if dur < 21 else 'wx_l'
        return 'wx_l'

tool.UCMProcessor._classify_seg = new_classify

uproc = tool.UCMProcessor({'dl_peak_limit': 900, 'ul_peak_limit': 160})
paths = [
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115328-电信.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115334-电信.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336-联通.xlsx'
]
params = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10, 'match_mode': 'auto'}
agg = uproc.parse_mmf(paths, params)
uproc._fn = 'mmf_test'

round_results = uproc._match_auto(agg, params, 50, 10)

from collections import Counter
key_counts = Counter()
for rd in round_results:
    for k in rd:
        key_counts[k] += 1

print(f"总轮次: {len(round_results)}")
print(f"各业务次数: {dict(key_counts)}")

# 检查ftp_dl数，应该~26
print(f"\n期望: store_l~26, store_s~26, ftp_dl~26, ftp_ul~26, wx_l~26, wx_s~26")