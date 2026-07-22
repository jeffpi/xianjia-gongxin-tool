#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
精确方案：合并后按dur+rate区分FTP高速下载 vs store_l
关键洞察：
- FTP下载高速段 flow=700-1500, dur=10-15s, rate=600-1100Mbps
- store_l 合并后 flow=600-1100, dur=15-60s, rate=200-500Mbps
所以：合并后，flow>=600且dur>=15且rate<500 → store_l
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

# 新分类器：合并后，flow>=600且(dur>=15或rate<500)→store_l
def new_classify(self, s, segs, idx):
    d = s['direction']; flow = s['flow_mb']; dur = s['duration']; rate = s['rate']
    if d == '下行':
        if flow <= 50:
            return None
        if 50 < flow < 350:
            return 'store_s'
        if 350 <= flow < 700:
            return 'ftp_dl'
        # flow >= 700
        # 关键改动：dur>=15 或 rate<500 → store_l（商店大包）
        # 否则 → ftp_dl（FTP高速下载段）
        if dur >= 15 or rate < 500:
            return 'store_l'
        else:
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

# 显示所有store_l详情
print("\n===== store_l 详情 =====")
for i, rd in enumerate(round_results):
    for k, s in rd.items():
        if k == 'store_l':
            print(f"  轮{i}: flow={s['flow_mb']:.0f}MB dur={s['duration']:.1f}s rate={s['rate']:.1f}Mbps "
                  f"start={s['start_t'].strftime('%H:%M:%S')}-{s['end_t'].strftime('%H:%M:%S')}")

print(f"\n===== ftp_dl 详情 =====")
for i, rd in enumerate(round_results):
    for k, s in rd.items():
        if k == 'ftp_dl':
            print(f"  轮{i}: flow={s['flow_mb']:.0f}MB dur={s['duration']:.1f}s rate={s['rate']:.1f}Mbps "
                  f"start={s['start_t'].strftime('%H:%M:%S')}")

print(f"\n===== store_s 详情 =====")
for i, rd in enumerate(round_results):
    for k, s in rd.items():
        if k == 'store_s':
            print(f"  轮{i}: flow={s['flow_mb']:.0f}MB dur={s['duration']:.1f}s rate={s['rate']:.1f}Mbps "
                  f"start={s['start_t'].strftime('%H:%M:%S')}")

print(f"\n期望: store_l~26, store_s~26, ftp_dl~26, ftp_ul~26, wx_l~26, wx_s~26")