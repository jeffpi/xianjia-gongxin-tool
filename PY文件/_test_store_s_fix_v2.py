#!/usr/bin/env python3
"""临时测试脚本V2：验证 store_s 时长修复（使用真实工具内部逻辑）"""

import sys, os, json, warnings, traceback, re
warnings.filterwarnings('ignore')
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')

# 从工具中提取 _parse_t
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

class TestProcessor:
    def __init__(self):
        self.seconds = None
        self._filename = 'test'

    def parse_mmf(self, filepath):
        # 读mmf
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
        df['_t'] = df[tcol].apply(_parse_t) if tcol else None
        df = df[df['_t'].notna()].copy()
        print(f'有效行: {len(df)}')

        # 秒级聚合
        df['sec'] = df['_t'].dt.floor('s')
        dl_mac = _select_col(df, '下行', 'MAC')
        ul_mac = _select_col(df, '上行', 'MAC')
        dl_rlc = _select_col(df, '下行', 'RLC')
        ul_rlc = _select_col(df, '上行', 'RLC')
        print(f'列: dl_mac={dl_mac}, ul_mac={ul_mac}, dl_rlc={dl_rlc}, ul_rlc={ul_rlc}')

        agg = df.groupby('sec').agg(
            dl_mac=(dl_mac, 'max'),
            ul_mac=(ul_mac, 'max'),
            dl_rlc=(dl_rlc, 'max'),
            ul_rlc=(ul_rlc, 'max'),
        ).reset_index().rename(columns={'sec': 't'})

        agg['dl_clip'] = agg['dl_rlc'].clip(upper=900)
        agg['ul_clip'] = agg['ul_rlc'].clip(upper=160)
        self.seconds = agg
        print(f'聚合后: {len(agg)} 行, dl_mac>0: {(agg["dl_mac"]>0).sum()}, dl_rlc>0: {(agg["dl_rlc"]>0).sum()}')
        return len(agg)

    def _classify_seg(self, s, segs, idx):
        d = s['direction']; flow = s['flow_mb']; dur = s['duration']; rate = s['rate']
        if d == '下行':
            if 50 < flow < 250:
                if dur > 20: return None  # 修复1
                return 'store_s'
            if flow > 700:
                if flow < 1500: return 'ftp_dl'
                return 'store_l'
            return None
        else:
            if flow < 20: return 'wx_s'
            if flow > 80:
                if flow > 190 and dur >= 21: return 'wx_l'
                if dur >= 21: return 'wx_l'
                if flow < 200: return 'ftp_ul'
                return 'wx_l'
            return None
        return None

    def _match_auto(self, agg, params, dl_min, ul_min):
        dl_min = dl_min or 50
        ul_min = ul_min or 10
        n = len(agg)
        segs = []
        i = 0
        while i < n:
            if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
                start_i = i; gap = 0
                while i < n and gap < 3:
                    if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
                        gap = 0
                    else:
                        gap += 1
                    i += 1
                end_i = max(start_i, i - 1 - gap)
                seg_df = agg.loc[start_i:end_i]
                ul_act = int((seg_df['ul_mac'] > 10).sum())
                dl_act = int((seg_df['dl_mac'] > 50).sum())
                if ul_act > dl_act:
                    direction, col, rlc_col, clip_col = '上行', 'ul_mac', 'ul_rlc', 'ul_clip'
                else:
                    direction, col, rlc_col, clip_col = '下行', 'dl_mac', 'dl_rlc', 'dl_clip'
                flow = seg_df[rlc_col].sum() / 8
                dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
                rv = seg_df[rlc_col].replace(0, np.nan).dropna()
                cv = seg_df[clip_col].replace(0, np.nan).dropna()
                segs.append({'start_i': start_i, 'end_i': end_i, 'start_t': seg_df['t'].iloc[0],
                             'end_t': seg_df['t'].iloc[-1], 'direction': direction,
                             'flow_mb': round(flow, 1), 'duration': round(dur, 1),
                             'rate': round(float(rv.mean()), 3) if len(rv) else 0,
                             'clip_rate': round(float(cv.mean()), 3) if len(cv) else 0,
                             'points': seg_df, 'result': '成功'})
            else:
                i += 1

        for idx, s in enumerate(segs):
            s['key'] = self._classify_seg(s, segs, idx)

        # 过滤
        key_segs = []
        for s in segs:
            k = s.get('key')
            if k is None: continue
            if k == 'store_s' and s.get('duration', 0) > 25:
                continue  # 修复2
            if k == 'wx_s':
                key_segs.append(s)
            elif s.get('duration', 0) >= 3:
                key_segs.append(s)

        rounds = []; cur_round = {}; last_end = None
        for s in key_segs:
            k = s['key']
            gap = (s['start_t'] - last_end).total_seconds() if last_end else 0
            if len(cur_round) > 0 and (gap > 120 or k in cur_round):
                if len(cur_round) >= 4 and 'ftp_dl' in cur_round and 'ftp_ul' in cur_round:
                    rounds.append(cur_round)
                cur_round = {}
            if k not in cur_round:
                cur_round[k] = s
                last_end = s['end_t']
        if len(cur_round) >= 4 and 'ftp_dl' in cur_round and 'ftp_ul' in cur_round:
            rounds.append(cur_round)

        return rounds

# ---------- 运行 ----------
proc = TestProcessor()
n_sec = proc.parse_mmf('/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx')
print(f'\n秒级聚合后: {n_sec} 行')

# auto 匹配
rounds = proc._match_auto(proc.seconds, {}, 50, 10)
print(f'\n总轮次: {len(rounds)}')

all_rates = {}
for k in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
    all_rates[k] = []

for r in rounds:
    for k, s in r.items():
        if k in all_rates:
            all_rates[k].append({'rate': s['clip_rate'], 'dur': s['duration'], 'flow': s['flow_mb']})

baseline = {
    'ftp_dl': 572.11,
    'ftp_ul': 66.39,
    'store_s': 176.29,
    'store_l': 1025.84,
    'wx_s': 27.81,
    'wx_l': 60.48,
}
key_name = {
    'ftp_dl': 'FTP下载', 'ftp_ul': 'FTP上传', 'store_s': '商店小包',
    'store_l': '商店大包', 'wx_s': '微信小文件', 'wx_l': '微信大文件',
}

print('\n=== 联通(store_s时长修复) ===')
for k, values in all_rates.items():
    n = len(values)
    if n == 0:
        print(f'{key_name[k]} N=0')
        continue
    rlc_avg = np.mean([v['rate'] for v in values])
    base = baseline[k]
    diff_pct = (rlc_avg - base) / base * 100
    flag = '✓' if abs(diff_pct) < 15 else ('±' if abs(diff_pct) < 25 else '✗')
    print(f'{key_name[k]} N={n} rate={rlc_avg:.2f} (基准{base}, {diff_pct:+.1f}%) {flag}')

print(f'\n总轮次={len(rounds)}')

print('\n=== store_s 段详情 ===')
for r_idx, r in enumerate(rounds):
    if 'store_s' in r:
        s = r['store_s']
        print(f'轮{r_idx+1}: dur={s["duration"]}s flow={s["flow_mb"]}MB rate={s["clip_rate"]:.1f}')

# 改回原始代码（不加修复2过滤）做对比
print('\n\n=== 对比：不加修复2的结果 ===')
proc2 = TestProcessor()
proc2.parse_mmf('/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx')

# 直接替换 _match_auto -- 不加store_s>25过滤
def _match_no_fix2(self2, agg, params, dl_min, ul_min):
    dl_min = dl_min or 50
    ul_min = ul_min or 10
    n = len(agg)
    segs = []
    i = 0
    while i < n:
        if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
            start_i = i; gap = 0
            while i < n and gap < 3:
                if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
                    gap = 0
                else:
                    gap += 1
                i += 1
            end_i = max(start_i, i - 1 - gap)
            seg_df = agg.loc[start_i:end_i]
            ul_act = int((seg_df['ul_mac'] > 10).sum())
            dl_act = int((seg_df['dl_mac'] > 50).sum())
            if ul_act > dl_act:
                direction, col, rlc_col, clip_col = '上行', 'ul_mac', 'ul_rlc', 'ul_clip'
            else:
                direction, col, rlc_col, clip_col = '下行', 'dl_mac', 'dl_rlc', 'dl_clip'
            flow = seg_df[rlc_col].sum() / 8
            dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
            rv = seg_df[rlc_col].replace(0, np.nan).dropna()
            cv = seg_df[clip_col].replace(0, np.nan).dropna()
            segs.append({'start_i': start_i, 'end_i': end_i, 'start_t': seg_df['t'].iloc[0],
                         'end_t': seg_df['t'].iloc[-1], 'direction': direction,
                         'flow_mb': round(flow, 1), 'duration': round(dur, 1),
                         'rate': round(float(rv.mean()), 3) if len(rv) else 0,
                         'clip_rate': round(float(cv.mean()), 3) if len(cv) else 0,
                         'points': seg_df, 'result': '成功'})
        else:
            i += 1

    for idx, s in enumerate(segs):
        # 用有修复1的_classify
        d = s['direction']; flow = s['flow_mb']; dur = s['duration']
        if d == '下行':
            if 50 < flow < 250:
                if dur > 20: return  # 修复1
                s['key'] = 'store_s'
            elif flow > 700:
                if flow < 1500: s['key'] = 'ftp_dl'
                else: s['key'] = 'store_l'
            else: s['key'] = None
        else:
            if flow < 20: s['key'] = 'wx_s'
            elif flow > 80:
                if dur >= 21: s['key'] = 'wx_l'
                elif flow < 200: s['key'] = 'ftp_ul'
                else: s['key'] = 'wx_l'
            else: s['key'] = None

    segs = [s for s in segs if s.get('key') is not None]
    # 不过滤store_s>25
    key_segs = []
    for s in segs:
        k = s['key']
        if k == 'wx_s':
            key_segs.append(s)
        elif s.get('duration', 0) >= 3:
            key_segs.append(s)

    rounds = []; cur_round = {}; last_end = None
    for s in key_segs:
        k = s['key']
        gap = (s['start_t'] - last_end).total_seconds() if last_end else 0
        if len(cur_round) > 0 and (gap > 120 or k in cur_round):
            if len(cur_round) >= 4 and 'ftp_dl' in cur_round and 'ftp_ul' in cur_round:
                rounds.append(cur_round)
            cur_round = {}
        if k not in cur_round:
            cur_round[k] = s
            last_end = s['end_t']
    if len(cur_round) >= 4 and 'ftp_dl' in cur_round and 'ftp_ul' in cur_round:
        rounds.append(cur_round)
    return rounds

rounds2 = _match_no_fix2(proc2, proc2.seconds, {}, 50, 10)
print(f'总轮次(无修复2): {len(rounds2)}')

all_rates2 = {}
for k in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
    all_rates2[k] = []
for r in rounds2:
    for k, s in r.items():
        if k in all_rates2:
            all_rates2[k].append({'rate': s['clip_rate'], 'dur': s['duration'], 'flow': s['flow_mb']})

print('\nstore_s段详情(无修复2):')
for r_idx, r in enumerate(rounds2):
    if 'store_s' in r:
        s = r['store_s']
        print(f'轮{r_idx+1}: dur={s["duration"]}s flow={s["flow_mb"]}MB rate={s["clip_rate"]:.1f}')

# 看所有原始段中那些被classify为store_s且durations>25的
print('\n\n=== 原始段中store_s且>25s的段 ===')
proc3 = TestProcessor()
proc3.parse_mmf('/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx')
agg = proc3.seconds
segs_all = []
i = 0; n = len(agg)
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
        direction = '下行' if dl_act >= ul_act else '上行'
        rlc_col = 'dl_rlc' if direction == '下行' else 'ul_rlc'
        flow = seg_df[rlc_col].sum() / 8
        dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
        segs_all.append({'dur': dur, 'flow': flow, 'direction': direction, 'start': seg_df['t'].iloc[0], 'end': seg_df['t'].iloc[-1]})
    else:
        i += 1

for s in segs_all:
    if s['direction'] == '下行' and 50 < s['flow'] < 250 and s['dur'] > 20:
        if s['dur'] > 25:
            print(f'  dur={s["dur"]}s flow={s["flow"]:.0f}MB {s["start"]} -> {s["end"]} ***** >25s *****')
        else:
            print(f'  dur={s["dur"]}s flow={s["flow"]:.0f}MB {s["start"]} -> {s["end"]}')