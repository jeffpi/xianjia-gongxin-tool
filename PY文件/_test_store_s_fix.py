#!/usr/bin/env python3
"""临时测试脚本：验证 store_s 时长修复"""

import sys, os, json, warnings, traceback
warnings.filterwarnings('ignore')
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

# 直接内联 UCMProcessor 关键代码（避免导入GUI依赖）
sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')

# ---------- 复制 UCMProcessor 核心方法 ----------
class TestProcessor:
    def __init__(self):
        self.seconds = None
        self._filename = 'test'

    def _fmt_t(self, dt_val):
        if isinstance(dt_val, pd.Timestamp):
            return dt_val.strftime('%Y-%m-%d %H:%M:%S')
        return str(dt_val)

    def parse_mmf(self, filepath):
        """解析mmf — 直接从工具代码复制"""
        # 自动探测表头行
        for h in [0, 1, 6, 7, 8]:
            try:
                tdf = pd.read_excel(filepath, header=h)
            except:
                continue
            cols = list(tdf.columns)
            time_cands = [c for c in cols if isinstance(c, str) and ('时间' in c or '采集时间' in c or 'Time' in c or 'Timestamp' in c)]
            col_count = len(cols)
            if col_count >= 50 and (time_cands or col_count >= 100):
                header_row = h
                break
        else:
            header_row = 7  # fallback

        raw = pd.read_excel(filepath, header=header_row)
        print(f'表头行={header_row}, 列数={len(raw.columns)}, 行数={len(raw)}')

        # 找时间列
        time_col = None
        for c in raw.columns:
            s = str(c).strip()
            if '时间' in s or '采集时间' in s or 'Timestamp' in s or 'Time' in s:
                time_col = c
                break
        if time_col is None:
            # 第一列通常是时间
            time_col = raw.columns[0]

        raw[time_col] = pd.to_datetime(raw[time_col], errors='coerce')
        raw = raw.dropna(subset=[time_col]).reset_index(drop=True)
        raw = raw.sort_values(time_col).reset_index(drop=True)

        # 找速率列
        dl_mac_col = self._find_col(raw, '下行', 'MAC')
        ul_mac_col = self._find_col(raw, '上行', 'MAC')
        dl_rlc_col = self._find_col(raw, '下行', 'RLC')
        ul_rlc_col = self._find_col(raw, '上行', 'RLC')

        print(f'列: dl_mac={dl_mac_col}, ul_mac={ul_mac_col}, dl_rlc={dl_rlc_col}, ul_rlc={ul_rlc_col}')

        # 自动识别单位 bps/Mbps
        # 取中位数判断
        sample = raw[dl_mac_col].dropna().median() if dl_mac_col else 0
        is_bps = sample > 1000000

        if is_bps:
            factor = 1000000
            print(f'单位: bps (中位数={sample})')
        else:
            factor = 1
            print(f'单位: Mbps (中位数={sample})')

        # 构建秒级聚合
        raw = raw.copy()
        if dl_mac_col:
            raw['dl_mac'] = pd.to_numeric(raw[dl_mac_col], errors='coerce') / factor
        else:
            raw['dl_mac'] = 0
        if ul_mac_col:
            raw['ul_mac'] = pd.to_numeric(raw[ul_mac_col], errors='coerce') / factor
        else:
            raw['ul_mac'] = 0
        if dl_rlc_col:
            raw['dl_rlc'] = pd.to_numeric(raw[dl_rlc_col], errors='coerce') / factor
        else:
            raw['dl_rlc'] = 0
        if ul_rlc_col:
            raw['ul_rlc'] = pd.to_numeric(raw[ul_rlc_col], errors='coerce') / factor
        else:
            raw['ul_rlc'] = 0

        raw['t'] = raw[time_col]

        # 秒级聚合: MAC取max, RLC取max
        raw['t_floor'] = raw['t'].dt.floor('1s')
        grp = raw.groupby('t_floor').agg({
            'dl_mac': 'max', 'ul_mac': 'max',
            'dl_rlc': 'max', 'ul_rlc': 'max',
            't': 'first'
        }).reset_index()

        # 削峰
        grp['dl_clip'] = grp['dl_rlc'].clip(upper=1000)
        grp['ul_clip'] = grp['ul_rlc'].clip(upper=200)

        self.seconds = grp.copy()
        return len(grp)

    def _find_col(self, df, direction, layer):
        dir_kw = ['下行', 'Downlink'] if direction == '下行' else ['上行', 'Uplink']
        cands = [c for c in df.columns if any(k in str(c) for k in dir_kw)
                 and layer in str(c) and ('吞吐率' in str(c) or 'hroughput' in str(c).lower())]
        return cands[0] if cands else None

    def _classify_seg(self, s, segs, idx):
        d = s['direction']; flow = s['flow_mb']; dur = s['duration']; rate = s['rate']
        if d == '下行':
            if 50 < flow < 250:
                if dur > 20: return None  # <-- 修复1
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
                continue  # <-- 修复2
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

# ---------- 运行测试 ----------
proc = TestProcessor()
n_sec = proc.parse_mmf('/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx')
print(f'\n秒级聚合后: {n_sec} 行')

# 运行 auto 匹配
rounds = proc._match_auto(proc.seconds, {}, 50, 10)
print(f'\n总轮次: {len(rounds)}')

# 统计各业务
all_rates = {}
for k in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
    all_rates[k] = []

for r in rounds:
    for k, s in r.items():
        if k in all_rates:
            all_rates[k].append({'rate': s['clip_rate'], 'dur': s['duration'], 'flow': s['flow_mb']})

# 基准
baseline = {
    'ftp_dl': 572.11,
    'ftp_ul': 66.39,
    'store_s': 176.29,
    'store_l': 1025.84,
    'wx_s': 27.81,
    'wx_l': 60.48,
}

key_name = {
    'ftp_dl': 'FTP下载',
    'ftp_ul': 'FTP上传',
    'store_s': '商店小包',
    'store_l': '商店大包',
    'wx_s': '微信小文件',
    'wx_l': '微信大文件',
}

print('\n=== 联通(store_s时长修复) ===')
for k, values in all_rates.items():
    n = len(values)
    if n == 0:
        print(f'{key_name[k]} N=0 (无数据)')
        continue
    rlc_avg = np.mean([v['rate'] for v in values])
    dur_avg = np.mean([v['dur'] for v in values])
    base = baseline[k]
    diff_pct = (rlc_avg - base) / base * 100
    flag = '✓' if abs(diff_pct) < 15 else ('±' if abs(diff_pct) < 25 else '✗')
    print(f'{key_name[k]} N={n} RLC={rlc_avg:.2f} (基准{base}, {diff_pct:+.1f}%) dur={dur_avg:.1f}s {flag}')

print(f'\n总轮次={len(rounds)}')

# 详细段信息
print('\n=== store_s 段详情 ===')
store_s_segs = []
for r_idx, r in enumerate(rounds):
    if 'store_s' in r:
        s = r['store_s']
        store_s_segs.append(f"轮{r_idx+1}: dur={s['duration']}s flow={s['flow_mb']}MB rate={s['clip_rate']:.1f} clip_rate={s['rate']:.1f}")
for line in store_s_segs:
    print(line)
