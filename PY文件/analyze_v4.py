#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""分析业务规律 - V4（按秒聚合+高速段检测）"""
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
from collections import Counter

df = pd.read_excel('联通/mmf20260703-电信-合并-带基准_V3.xlsx')

def parse_time(t):
    try:
        t = str(t).strip()
        m = re.match(r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\((\d+)\)', t)
        if m:
            dt = datetime.strptime(m.group(1), '%Y-%m-%d %H:%M:%S')
            ms = int(m.group(2))
            return dt + timedelta(milliseconds=ms)
        return None
    except:
        return None

df['_c_time'] = df['Time'].apply(parse_time)
qci_col = 'QCI'
dl_rlc_col = [c for c in df.columns if 'Downlink RLC Throughput' in c][0]
ul_rlc_col = [c for c in df.columns if 'Uplink RLC Throughput' in c][0]

df['_dl_mb'] = df[dl_rlc_col] / 1e6   # Mbps
df['_ul_mb'] = df[ul_rlc_col] / 1e6

# 只看QCI=7
df7 = df[df[qci_col] == 7].copy().sort_values('_c_time').reset_index(drop=True)

# 按秒聚合：同一秒的多条记录，DL/UL取max
df7['_sec'] = df7['_c_time'].dt.floor('s')
sec_agg = df7.groupby('_sec').agg({
    '_dl_mb': 'max',
    '_ul_mb': 'max',
}).reset_index()

print(f"QCI=7原始行数: {len(df7)}")
print(f"按秒聚合后: {len(sec_agg)} 秒")

dl_sec = sec_agg['_dl_mb'].fillna(0).values
ul_sec = sec_agg['_ul_mb'].fillna(0).values
ts_sec = sec_agg['_sec'].values

# 找高速段：DL>50或UL>20算活跃
DL_ACTIVE = 50.0  # Mbps
UL_ACTIVE = 20.0  # Mbps

active = (dl_sec > DL_ACTIVE) | (ul_sec > UL_ACTIVE)

# 合并相邻活跃秒（间隔<=2秒的算同一段）
segments = []
i = 0
n = len(sec_agg)
while i < n:
    if active[i]:
        si = i; st = ts_sec[i]; j = i + 1
        while j < n:
            # 检查是否有间隙>2秒的非活跃
            if not active[j]:
                # 看看后面2秒内是否有活跃
                k = j
                while k < n and not active[k] and (ts_sec[k] - ts_sec[j-1]) / np.timedelta64(1,'s') <= 2:
                    k += 1
                if k < n and active[k]:
                    j = k
                    continue
                else:
                    break
            j += 1
        et = ts_sec[j-1]; dur = (et - st) / np.timedelta64(1, 's')
        dl_sum = float(dl_sec[si:j].sum()); ul_sum = float(ul_sec[si:j].sum())
        segments.append({'st': st, 'et': et, 'dur': dur + 1, 'dl': dl_sum, 'ul': ul_sum})
        i = j
    else:
        i += 1

print(f'高速段: {len(segments)} 个')

# 识别业务
biz = []
for seg in segments:
    dur = seg['dur']; dl_s = seg['dl']; ul_s = seg['ul']
    bt = None

    # FTP: 时长8.5-11.5秒
    if 8.5 <= dur <= 11.5:
        if dl_s > ul_s:
            if 8800 <= dl_s <= 10100:
                bt = '应用商店大文件下载'
            else:
                bt = 'FTP下载'
        else:
            bt = 'FTP上传'
    # 应用商店大文件: dl在8800-10100, 时长>6
    elif dl_s > ul_s and 8800 <= dl_s <= 10100 and dur > 6:
        bt = '应用商店大文件下载'
    # 应用商店小文件: 时长2.5-7秒, dl 100-900
    elif 2.5 <= dur <= 7.5 and dl_s > ul_s and 100 <= dl_s <= 900:
        bt = '应用商店小文件下载'
    # 微信小包: 时长0.5-3秒, ul 20-200
    elif 0.5 <= dur <= 3.5 and ul_s > dl_s and 20 <= ul_s <= 200:
        bt = '微信小包发送'
    # 微信大包: 时长15-55秒, ul>=800
    elif 15 <= dur <= 60 and ul_s > dl_s and ul_s >= 800:
        bt = '微信大包发送'

    if bt:
        seg['bt'] = bt
        biz.append(seg)

print(f'识别业务: {len(biz)} 个')

for i, seg in enumerate(biz[:60], 1):
    st_str = pd.Timestamp(seg['st']).strftime('%H:%M:%S')
    print(f'{i:3d}. {seg["bt"]:12s} | {st_str} | {seg["dur"]:6.2f}s | DL:{seg["dl"]:8.2f} UL:{seg["ul"]:8.2f} | ÷8: DL={seg["dl"]/8:.1f}MB UL={seg["ul"]/8:.1f}MB')

cnt = Counter(s['bt'] for s in biz)
print(f'\n业务统计:')
for k, v in sorted(cnt.items()):
    print(f'  {k}: {v} 个')

# 标记G列（将业务标记写回原始df）
# 需要把秒级段映射回原始行
df['业务标记'] = ''
for seg in biz:
    st = seg['st']; et = seg['et']
    # 找到原始df中在该时间段内且QCI=7的行
    mask = (df['_c_time'] >= st) & (df['_c_time'] <= et + timedelta(seconds=1)) & (df[qci_col] == 7)
    df.loc[mask, '业务标记'] = seg['bt']

# 重新排列列
cols = df.columns.tolist()
ti = cols.index('Time')
new_cols = cols[:ti+1] + ['下载速率', '上传速率', '持续时长', '临时项', '业务标记'] + [c for c in cols if c not in set(cols[:ti+1] + ['下载速率', '上传速率', '持续时长', '临时项', '业务标记', '_c_time', '_dl_mb', '_ul_mb'])]
df = df[new_cols]

output_file = '联通/mmf20260703-电信-合并-带基准_V4.xlsx'
df.to_excel(output_file, index=False)
print(f'\n已保存: {output_file}')