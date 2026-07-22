#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
精确方案：threshold=600且store_l合并后需要加额外按时间位置区分的逻辑。
回退：先确认不合并时ftp_dl个数的变化，然后只调合并+阈值
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

# 先看原始分类+合并后（即原gap<5 merge），store_l 和 ftp_dl的分布
uproc = tool.UCMProcessor({'dl_peak_limit': 900, 'ul_peak_limit': 160})
paths = [
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115328-电信.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115334-电信.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336-联通.xlsx'
]
params = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10, 'match_mode': 'auto'}
agg = uproc.parse_mmf(paths, params)

# 提取原始段（不合并）
n = len(agg)
segs = []
i = 0
dl_min, ul_min = 50, 10
while i < n:
    if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
        start_i = i; gap = 0
        while i < n and gap < 5:
            if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
                gap = 0
            else:
                gap += 1
            i += 1
        end_i = max(start_i, i - 1 - gap)
        seg_df = agg.loc[start_i:end_i]
        ul_act = int((seg_df['ul_mac'] > 10).sum())
        dl_act = int((seg_df['dl_mac'] > 50).sum())
        if ul_act > dl_act + 3:
            direction, col = '上行', 'ul_mac'
        else:
            direction, col = '下行', 'dl_mac'
        flow = seg_df[col].sum() / 8
        dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
        segs.append({'direction': direction, 'flow_mb': round(flow,1), 'duration': round(dur,1)})
    else:
        i += 1

# 看下行段无合并的flow分布
print("===== 下行段（gap=5切割，无合并）=====")
dl_segs = [s for s in segs if s['direction'] == '下行']
dl_segs = sorted(dl_segs, key=lambda x: x['flow_mb'])
for s in dl_segs:
    f = s['flow_mb']
    if f > 100:  # 只看有意义的
        k = 'store_s' if 50<f<350 else ('ftp_dl' if f<700 else ('ftp_dl' if f<1500 else 'store_l'))
        print(f"  flow={f:6.0f}MB dur={s['duration']:5.1f}s -> {k}")

print(f"下行段总数: {len(dl_segs)}")
print(f"flow>=600: {len([s for s in dl_segs if s['flow_mb']>=600])}")
print(f"flow>=700: {len([s for s in dl_segs if s['flow_mb']>=700])}")
print(f"flow<700: {len([s for s in dl_segs if s['flow_mb']<700])}")
print(f"flow 在600-700: {len([s for s in dl_segs if 600 <= s['flow_mb'] < 700])}")
print(f"flow 在700-1500: {len([s for s in dl_segs if 700 <= s['flow_mb'] < 1500])}")

# 现在看看：如果threshold=700，哪些FTP会被误判成store_l？
print("\n===== 下行flow在700-1500(原本是FTP，如果降到700会被误判) =====")
mid = [s for s in dl_segs if 700 <= s['flow_mb'] < 1500 and s['duration'] < 20]
print(f"flow 700-1500且dur<20s: {len(mid)}个 — 这些很可能是FTP高速段")
for s in mid:
    print(f"  flow={s['flow_mb']:.0f}MB dur={s['duration']:.1f}s — 速率={s['flow_mb']*8/s['duration']:.0f}Mbps")

mid2 = [s for s in dl_segs if 700 <= s['flow_mb'] < 1500 and s['duration'] >= 15]
print(f"\nflow 700-1500且dur>=15s: {len(mid2)}个 — 这些可能是真正的store_l")
for s in mid2:
    print(f"  flow={s['flow_mb']:.0f}MB dur={s['duration']:.1f}s — 速率={s['flow_mb']*8/s['duration']:.0f}Mbps")
