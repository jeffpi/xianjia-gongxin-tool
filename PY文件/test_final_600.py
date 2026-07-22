#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最终方案：将threshold降到600 + 不改变FTP下载/上传的检测
关键调整：ftp_dl threshold从700降到600，维持dur+rate区分
"""

import numpy as np
import pandas as pd
from importlib import util as import_util

spec = import_util.spec_from_file_location('tool', '5G用户级公共监控速率统计工具-V1.04.py')
tool = import_util.module_from_spec(spec)
spec.loader.exec_module(tool)

# 最终版本分类器：threshold降到600
def final_classify(self, s, segs, idx):
    d = s['direction']; flow = s['flow_mb']; dur = s['duration']; rate = s['rate']
    if d == '下行':
        if flow <= 50:
            return None
        if 50 < flow < 350:
            return 'store_s'
        if 350 <= flow < 600:
            return 'ftp_dl'
        # flow >= 600
        if flow >= 1500:
            return 'store_l'
        # 600~1500: 区分FTP高速下载(短时高rate) vs store_l(长时低rate)
        if dur >= 15:
            return 'store_l'
        if rate < 500:
            return 'store_l'
        return 'ftp_dl'  # 短时高速 → FTP
    else:
        if flow < 20:
            return 'wx_s'
        if 20 <= flow <= 80:
            return 'ftp_ul'
        if 80 < flow < 200:
            return 'ftp_ul' if dur < 21 else 'wx_l'
        if 200 <= flow < 300:
            return 'ftp_ul' if (rate > 70 and dur < 15) else 'wx_l'
        return 'wx_l'

tool.UCMProcessor._classify_seg = final_classify

uproc = tool.UCMProcessor({'dl_peak_limit': 900, 'ul_peak_limit': 160})
paths = ['联通/mmf20260703115328-电信.xlsx','联通/mmf20260703115334-电信.xlsx','联通/mmf20260703115336-联通.xlsx']
params = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10, 'match_mode': 'auto'}
agg = uproc.parse_mmf(paths, params)
uproc._fn = 'test'

round_results = uproc._match_auto(agg, params, 50, 10)

from collections import Counter
key_counts = Counter()
for rd in round_results:
    for k in rd:
        key_counts[k] += 1
print(f"总轮次: {len(round_results)}")
print(f"各业务次数: {dict(key_counts)}")
print(f"\n期望: store_l~26, store_s~26, ftp_dl~26, ftp_ul~26, wx_l~26")
print(f"实际: store_l={key_counts.get('store_l')}, store_s={key_counts.get('store_s')}, ftp_dl={key_counts.get('ftp_dl')}, ftp_ul={key_counts.get('ftp_ul')}")

# 统计ftp_dl flow
ftp_dl_flows = [s['flow_mb'] for rd in round_results for k, s in rd.items() if k == 'ftp_dl']
if ftp_dl_flows:
    arr = np.array(ftp_dl_flows)
    print(f"\nftp_dl: min={arr.min():.0f} max={arr.max():.0f} mean={arr.mean():.0f} n={len(arr)}")
    print(f"  >=600: {sum(arr>=600)}, >=700: {sum(arr>=700)}, >=800: {sum(arr>=800)}")

# 统计store_l flow
store_l_flows = [s['flow_mb'] for rd in round_results for k, s in rd.items() if k == 'store_l']
if store_l_flows:
    arr = np.array(store_l_flows)
    print(f"store_l: min={arr.min():.0f} max={arr.max():.0f} mean={arr.mean():.0f} n={len(arr)}")
    print(f"  <=800: {sum(arr<=800)}, <=1000: {sum(arr<=1000)}, >5000: {sum(arr>5000)}")

# 检查 store_s
store_s_flows = [s['flow_mb'] for rd in round_results for k, s in rd.items() if k == 'store_s']
if store_s_flows:
    arr = np.array(store_s_flows)
    print(f"store_s: min={arr.min():.0f} max={arr.max():.0f} mean={arr.mean():.0f} n={len(arr)}")