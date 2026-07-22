#!/usr/bin/env python3
"""诊断脚本：检查原始段分类和轮次组成"""

import sys, os, warnings, re
warnings.filterwarnings('ignore')
import numpy as np
import pandas as pd

sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')

with open('/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py', 'r', encoding='utf-8') as f:
    source = f.read()

lines = source.split('\n')
ucm_start = None
for i, line in enumerate(lines):
    if line.startswith('class UCMProcessor'):
        ucm_start = i
        break

imports_code = '\n'.join(lines[:ucm_start])
ucm_code = '\n'.join(lines[ucm_start:])
ucm_code = ucm_code.replace('self._filename', '"test"')

exec(imports_code + '\n' + ucm_code)

# 运行
config = {"dl_peak_limit": 900, "ul_peak_limit": 160, "down_min": 50, "up_min": 10}
proc = UCMProcessor(config)
agg = proc.parse_mmf("/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx", config)
params = {"match_mode": "auto", "down_min": 50, "up_min": 10}

# 获取所有原始段
n = len(agg)
segs_all = []
i = 0
while i < n:
    if agg.at[i, 'dl_mac'] > 50 or agg.at[i, 'ul_mac'] > 10:
        start_i = i; gap = 0
        while i < n and gap < 3:
            if agg.at[i, 'dl_mac'] > 50 or agg.at[i, 'ul_mac'] > 10:
                gap = 0
            else:
                gap += 1
            i += 1
        end_i = max(start_i, i - 1 - gap)
        seg_df = agg.loc[start_i:end_i]
        ul_act = int((seg_df['ul_mac'] > 10).sum())
        dl_act = int((seg_df['dl_mac'] > 50).sum())
        direction = '上行' if ul_act > dl_act else '下行'
        rlc_col = 'ul_rlc' if direction == '上行' else 'dl_rlc'
        flow = seg_df[rlc_col].sum() / 8
        dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
        rv = seg_df[rlc_col].replace(0, np.nan).dropna()
        rate = round(float(rv.mean()), 3) if len(rv) else 0
        segs_all.append({
            'start': seg_df['t'].iloc[0], 'end': seg_df['t'].iloc[-1],
            'dur': dur, 'flow': round(flow,1), 'direction': direction, 'rate': rate,
            'dl_rlc_sum': seg_df['dl_rlc'].sum() / 8,
            'ul_rlc_sum': seg_df['ul_rlc'].sum() / 8,
        })
    else:
        i += 1

# 分类（用修复后的规则）
for s in segs_all:
    if s['direction'] == '下行':
        if 50 < s['flow'] < 250:
            if s['dur'] > 20:
                s['key'] = None  # 被拦截
            else:
                s['key'] = 'store_s'
        elif s['flow'] > 700:
            if s['flow'] < 1500: s['key'] = 'ftp_dl'
            else: s['key'] = 'store_l'
        else: s['key'] = None
    else:
        if s['flow'] < 20: s['key'] = 'wx_s'
        elif s['flow'] > 80:
            if s['dur'] >= 21: s['key'] = 'wx_l'
            elif s['flow'] < 200: s['key'] = 'ftp_ul'
            else: s['key'] = 'wx_l'
        else: s['key'] = None

# 过滤进key_segs
key_segs = []
for s in segs_all:
    k = s.get('key')
    if k is None: continue
    if k == 'store_s' and s['dur'] > 25:
        s['_filtered'] = 'store_s>25'
        continue
    if k == 'wx_s':
        key_segs.append(s)
    elif s['dur'] >= 3:
        key_segs.append(s)
    else:
        s['_filtered'] = 'dur<3'

# 分轮
rounds = []
cur_round = {}
last_end = None
for s in key_segs:
    k = s['key']
    gap = (s['start'] - last_end).total_seconds() if last_end else 0
    if len(cur_round) > 0 and (gap > 120 or k in cur_round):
        if len(cur_round) >= 4 and 'ftp_dl' in cur_round and 'ftp_ul' in cur_round:
            rounds.append(cur_round)
        cur_round = {}
    if k not in cur_round:
        cur_round[k] = s
        last_end = s['end']
if len(cur_round) >= 4 and 'ftp_dl' in cur_round and 'ftp_ul' in cur_round:
    rounds.append(cur_round)

print(f"总原始段: {len(segs_all)}")
print(f"进key_segs: {len(key_segs)}")
print(f"总轮次: {len(rounds)}")

print("\n=== 被过滤的段 ===")
for s in segs_all:
    if s.get('_filtered'):
        print(f"  {s['_filtered']}: {s.get('key','')} dur={s['dur']}s flow={s['flow']}MB {s['start']}")

print("\n=== 各轮key组成 ===")
for r_idx, r in enumerate(rounds):
    parts = []
    for k, s in r.items():
        parts.append(f"{k}({s['dur']}s)")
    print(f"  轮{r_idx+1}: {', '.join(parts)}")

print("\n=== store_s段详情 ===")
for r_idx, r in enumerate(rounds):
    if 'store_s' in r:
        s = r['store_s']
        print(f"  轮{r_idx+1}: dur={s['dur']}s flow={s['flow']}MB rate={s['rate']:.1f} {s['start']}")

# 统计rate
print("\n=== 各业务速率 ===")
baseline = {'ftp_dl': 572.11, 'ftp_ul': 66.39, 'store_s': 176.29, 'store_l': 1025.84, 'wx_s': 27.81, 'wx_l': 60.48}
key_name = {'ftp_dl': 'FTP下载', 'ftp_ul': 'FTP上传', 'store_s': '商店小包', 'store_l': '商店大包', 'wx_s': '微信小文件', 'wx_l': '微信大文件'}

for k in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
    values = [r[k]['rate'] for r in rounds if k in r]
    n = len(values)
    if n == 0:
        print(f"{key_name[k]} N=0")
        continue
    rlc_avg = np.mean(values)
    base = baseline[k]
    diff_pct = (rlc_avg - base) / base * 100
    flag = 'V' if abs(diff_pct) < 15 else ('-' if abs(diff_pct) < 25 else 'X')
    print(f"{key_name[k]} N={n} RLC={rlc_avg:.2f} (基准{base}, {diff_pct:+.1f}%) {flag}")