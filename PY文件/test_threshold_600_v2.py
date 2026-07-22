#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试：降低threshold=600，检查结果
原因：SL#1的841MB/12s段rate=747,dur<15→被分ftp_dl，但它是store_l
SL#2的740MB/9s同样
这些段的flow在700~1500之间，但因为dur<15且rate>=500而被分ftp_dl
如果threshold降到600，flow>=600且dur<15且rate>=500的段会变成什么？
"""

import numpy as np
import pandas as pd
from importlib import util as import_util

spec = import_util.spec_from_file_location('tool', '5G用户级公共监控速率统计工具-V1.04.py')
tool = import_util.module_from_spec(spec)
spec.loader.exec_module(tool)

# 新分类：flow>=600就检查，而不只是flow>=700
def new_classify(self, s, segs, idx):
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
        # 600~1500
        if dur >= 15:
            return 'store_l'
        if rate < 500:
            return 'store_l'
        return 'ftp_dl'
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

tool.UCMProcessor._classify_seg = new_classify

uproc = tool.UCMProcessor({'dl_peak_limit': 900, 'ul_peak_limit': 160})
paths = ['联通/mmf20260703115328-电信.xlsx','联通/mmf20260703115334-电信.xlsx','联通/mmf20260703115336-联通.xlsx']
params = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10, 'match_mode': 'auto'}
agg = uproc.parse_mmf(paths, params)

from collections import Counter

# 直接跑_match_auto
uproc._fn = 'mmf_test'
round_results = uproc._match_auto(agg, params, 50, 10)
key_counts = Counter()
for rd in round_results:
    for k in rd:
        key_counts[k] += 1
print(f"总轮次: {len(round_results)}")
print(f"各业务次数: {dict(key_counts)}")

# 查看ftp_dl的flow分布
print("\n===== ftp_dl flow分布 =====")
from collections import defaultdict
ftp_dl_flows = defaultdict(list)
for rd in round_results:
    for k, s in rd.items():
        if k == 'ftp_dl':
            ftp_dl_flows[s['flow_mb']].append(s)
# 按flow排序
for flow in sorted(ftp_dl_flows.keys()):
    segs = ftp_dl_flows[flow]
    print(f"  flow={flow:.0f}MB: {len(segs)}个")
    for s in segs:
        print(f"    dur={s['duration']:.1f}s rate={s['rate']:.1f}")

# 统计flow区间
ftp_dl_flow_list = []
ftp_ul_flow_list = []
for rd in round_results:
    for k, s in rd.items():
        if k == 'ftp_dl':
            ftp_dl_flow_list.append(s['flow_mb'])
        elif k == 'ftp_ul':
            ftp_ul_flow_list.append(s['flow_mb'])

if ftp_dl_flow_list:
    arr = np.array(ftp_dl_flow_list)
    print(f"\nftp_dl flow: min={arr.min():.0f} max={arr.max():.0f} mean={arr.mean():.0f}")
    print(f"  >=600: {np.sum(arr>=600)}个")
    print(f"  >=700: {np.sum(arr>=700)}个")
    print(f"  >=800: {np.sum(arr>=800)}个")