#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
方案2：合并后按flow+dur+rate分类
关键发现：
- store_l在参考文件中的分段是10~30s短高速段，flow=572~1102MB
- 这些段rate通常在800~1100Mbps
- FTP下载高速段也是dur=10s, rate同样高

但store_l是**连续高速下行**，而FTP下载总flow较小(300-700)，偶尔也到700-1400

更好的区分方式：看这些段的**开始时间**在轮次中的位置
- store_l 总是紧跟在 store_s 之后（同一轮）
- ftp_dl 在轮次最前面

但实际上我们不知道它们所属轮次，需要先分类...

回退方案：利用段内数据的**分布特征**
- store_l: 持续高速，RLC rate稳定在较高水平
- FTP下载高速段：但flow更大的FTP下载段，可能被当作store_l

关键洞察：参考文件中store_l的总flow在MAC口径下是572~1102MB
而现在是gap=5切割+gap<10合并后的flow已经足够大

实测发现：目前store_l有22个，还差4个。
漏了哪些？轮6(flow=812MB, dur=22s)应该在参考中，但被分类成ftp_dl？
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

# 新分类器：flow>=700且dur>=15 → store_l; flow>=1500 → store_l; 剩余→ftp_dl
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
        # store_l: 长时段(>=15s)或flow>=1500
        if flow >= 1500:
            return 'store_l'
        if dur >= 15:
            return 'store_l'
        # 短段(10-14s)且rate高 → 可能是FTP下载高速
        if rate > 500:
            return 'ftp_dl'  # 高速FTP下载
        # 短段但rate低 → 可能是差信号的store_l
        return 'store_l'
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