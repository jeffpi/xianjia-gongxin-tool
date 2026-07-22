#!/usr/bin/env python3
"""诊断：看原始段中store_l和store_s的分类情况"""

import sys, os, warnings, re
warnings.filterwarnings('ignore')
import numpy as np
import pandas as pd

def _parse_t(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ('', 'nan', 'None'):
        return None
    try:
        return pd.to_datetime(re.sub(r'\(\d+\)', '', str(v)).strip())
    except Exception:
        return None

def _find_col(df, keywords):
    for c in df.columns:
        cs = str(c)
        if any(k in cs for k in keywords):
            return c
    return None

def _select_col(df, direction, layer):
    dir_kw = ['下行', 'Downlink'] if direction == '下行' else ['上行', 'Uplink']
    cands = [c for c in df.columns if any(k in str(c) for k in dir_kw)
             and layer in str(c) and ('吞吐率' in str(c) or 'hroughput' in str(c).lower())]
    return cands[0] if cands else None

# 读取
filepath = '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx'
raw = pd.read_excel(filepath, header=None, nrows=10)
header_row = 0
for i in range(min(10, len(raw))):
    vals = [str(v) for v in raw.iloc[i].tolist()]
    if any('吞吐率' in v or 'hroughput' in v.lower() for v in vals):
        header_row = i; break

df = pd.read_excel(filepath, header=header_row)
df = df.replace('-', np.nan)
for col in list(df.columns):
    cs = str(col); low = cs.lower()
    if ('吞吐率' in cs or 'hroughput' in low) and 'bps' in low and 'mbps' not in low:
        new_col = re.sub(r'\(.*?bps.*?\)', '(Mbps)', cs, flags=re.IGNORECASE)
        if new_col == cs: new_col = cs + '(Mbps)'
        if new_col not in df.columns:
            df[new_col] = pd.to_numeric(df[col], errors='coerce') / 1_000_000
        df = df.drop(columns=[col])

tcol = _find_col(df, ['采集时间', 'Time'])
df['_t'] = df[tcol].apply(_parse_t)
df = df[df['_t'].notna()].copy()
df['sec'] = df['_t'].dt.floor('s')
dl_mac = _select_col(df, '下行', 'MAC')
ul_mac = _select_col(df, '上行', 'MAC')
dl_rlc = _select_col(df, '下行', 'RLC')
ul_rlc = _select_col(df, '上行', 'RLC')

agg = df.groupby('sec').agg(
    dl_mac=(dl_mac, 'max'),
    ul_mac=(ul_mac, 'max'),
    dl_rlc=(dl_rlc, 'max'),
    ul_rlc=(ul_rlc, 'max'),
).reset_index().rename(columns={'sec': 't'})

# 切所有段
n = len(agg)
all_segs = []
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
        if ul_act > dl_act:
            direction, rlc_col = '上行', 'ul_rlc'
        else:
            direction, rlc_col = '下行', 'dl_rlc'
        flow = seg_df[rlc_col].sum() / 8
        dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
        rv = seg_df[rlc_col].replace(0, np.nan).dropna()
        rate = round(float(rv.mean()), 3) if len(rv) else 0

        # 分类
        key = None
        if direction == '下行':
            if 50 < flow < 250:
                if dur > 20:
                    key = f'store_s(DUR>{20})'  # 被拦截
                else:
                    key = 'store_s'
            elif flow > 700:
                if flow < 1500:
                    key = 'ftp_dl'
                else:
                    key = 'store_l'
        else:
            if flow < 20:
                key = 'wx_s'
            elif flow > 80:
                if dur >= 21:
                    key = 'wx_l'
                elif flow < 200:
                    key = 'ftp_ul'
                else:
                    key = 'wx_l'

        if key:
            all_segs.append({'key': key, 'dur': dur, 'flow': round(flow,1), 'rate': rate,
                            'direction': direction, 'start': seg_df['t'].iloc[0], 'end': seg_df['t'].iloc[-1]})
    else:
        i += 1

print(f'总段数(有key): {len(all_segs)}')
print(f'\n=== 各key数量 ===')
from collections import Counter
c = Counter(s['key'] for s in all_segs)
for k, v in sorted(c.items()):
    print(f'  {k}: {v}')

print(f'\n=== store_l 段详情 ===')
for s in all_segs:
    if s['key'] == 'store_l':
        print(f'  dur={s["dur"]}s flow={s["flow"]}MB rate={s["rate"]} {s["start"]}')

print(f'\n=== 下行flow>700的段 ===')
for s in all_segs:
    if s['direction'] == '下行' and s['flow'] > 700:
        print(f'  {s["key"]} dur={s["dur"]}s flow={s["flow"]}MB rate={s["rate"]} {s["start"]}')

print(f'\n=== store_s被拦截(dur>20)的段 ===')
for s in all_segs:
    if 'DUR' in s['key']:
        print(f'  dur={s["dur"]}s flow={s["flow"]}MB rate={s["rate"]} {s["start"]} -> {s["end"]}')

print(f'\n=== 下行段中flow在250~700之间的段 ===')
for s in all_segs:
    if s['direction'] == '下行' and 250 <= s['flow'] <= 700:
        print(f'  key={s["key"]} dur={s["dur"]}s flow={s["flow"]}MB rate={s["rate"]} {s["start"]}')