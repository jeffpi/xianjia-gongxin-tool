#!/usr/bin/env python3
"""
auto模式段识别各环节流失诊断脚本
不修改任何代码文件，纯分析
"""
import sys, os, re, datetime
import pandas as pd
import numpy as np
import openpyxl
from collections import Counter, defaultdict

TOOL_DIR = "/Users/sun/ClaudeCode/先甲工信部工具"
MMF_PATH = os.path.join(TOOL_DIR, "联通/mmf20260703115336-联通.xlsx")
REF_PATH = os.path.join(TOOL_DIR, "联通/工信部业务指标V2120260428_202607030329.xlsx")

# ========== Step 0: 复用工具的MMF解析 + 秒级聚合 ==========
def parse_t(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ('', 'nan', 'None'):
        return None
    try:
        return pd.to_datetime(re.sub(r'\(\d+\)', '', str(v)).strip())
    except:
        return None

def select_col(df, direction, layer):
    dir_kw = ['下行', 'Downlink'] if direction == '下行' else ['上行', 'Uplink']
    cands = [c for c in df.columns if any(k in str(c) for k in dir_kw)
             and layer in str(c) and ('吞吐率' in str(c) or 'hroughput' in str(c).lower())]
    return cands[0] if cands else None

def find_col(df, keywords):
    for c in df.columns:
        cs = str(c)
        if any(k in cs for k in keywords):
            return c
    return None

def load_mmf(path):
    raw = pd.read_excel(path, header=None, nrows=10)
    header_row = 0
    for i in range(min(10, len(raw))):
        vals = [str(v) for v in raw.iloc[i].tolist()]
        if any('吞吐率' in v or 'hroughput' in v.lower() for v in vals):
            header_row = i; break
    df = pd.read_excel(path, header=header_row)
    df = df.replace('-', np.nan)
    for col in list(df.columns):
        cs = str(col); low = cs.lower()
        if ('吞吐率' in cs or 'hroughput' in low) and 'bps' in low and 'mbps' not in low:
            new_col = re.sub(r'\(.*?bps.*?\)', '(Mbps)', cs, flags=re.IGNORECASE)
            if new_col == cs: new_col = cs + '(Mbps)'
            if new_col not in df.columns:
                df[new_col] = pd.to_numeric(df[col], errors='coerce') / 1_000_000
            df = df.drop(columns=[col])
    tcol = find_col(df, ['采集时间', 'Time'])
    df['_t'] = df[tcol].apply(parse_t) if tcol else None
    df = df[df['_t'].notna()].copy()
    return df

def aggregate(df, params=None):
    if params is None:
        params = {'dl_peak_limit': 900, 'ul_peak_limit': 160}
    df = df.sort_values('_t').reset_index(drop=True)
    df['sec'] = df['_t'].dt.floor('s')
    dl_mac = select_col(df, '下行', 'MAC'); ul_mac = select_col(df, '上行', 'MAC')
    dl_rlc = select_col(df, '下行', 'RLC'); ul_rlc = select_col(df, '上行', 'RLC')
    agg_d = {
        'dl_mac': (dl_mac, lambda s: pd.to_numeric(s, errors='coerce').max()),
        'ul_mac': (ul_mac, lambda s: pd.to_numeric(s, errors='coerce').max()),
        'dl_rlc': (dl_rlc, lambda s: pd.to_numeric(s, errors='coerce').max()),
        'ul_rlc': (ul_rlc, lambda s: pd.to_numeric(s, errors='coerce').mean()),
    }
    agg = df.groupby('sec').agg(**agg_d).reset_index().rename(columns={'sec': 't'})
    agg['dl_clip'] = agg['dl_rlc'].clip(upper=params.get('dl_peak_limit', 900))
    agg['ul_clip'] = agg['ul_rlc'].clip(upper=params.get('ul_peak_limit', 160))
    agg[['dl_mac', 'ul_mac', 'dl_rlc', 'ul_rlc']] = agg[['dl_mac', 'ul_mac', 'dl_rlc', 'ul_rlc']].fillna(0)
    return agg

# ========== Step 1: 段识别 (复制_match_auto第950-984行逻辑) ==========
def detect_segments(agg, dl_min=50, ul_min=10, gap_max=10):
    n = len(agg)
    segs = []
    i = 0
    while i < n:
        if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
            start_i = i; gap = 0
            while i < n and gap < gap_max:
                if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
                    gap = 0
                else:
                    gap += 1
                i += 1
            end_i = max(start_i, i - 1 - gap)
            seg_df = agg.loc[start_i:end_i]
            seg_dur = (seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds()
            ul_flow = float(seg_df['ul_mac'].sum()) / 8
            dl_flow = float(seg_df['dl_mac'].sum()) / 8
            ul_peak = float(seg_df['ul_mac'].max())
            dl_peak = float(seg_df['dl_mac'].max())
            # 方向判定
            if ul_flow > dl_flow * 0.5 or (seg_dur <= 3 and ul_peak > 15 and dl_peak < ul_peak):
                direction = '上行'; col = 'ul_mac'
            else:
                direction = '下行'; col = 'dl_mac'
            flow = seg_df[col].sum() / 8
            dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
            segs.append({
                'start_i': start_i, 'end_i': end_i,
                'start_t': seg_df['t'].iloc[0], 'end_t': seg_df['t'].iloc[-1],
                'direction': direction,
                'flow_mb': round(flow, 1), 'duration': round(dur, 1),
                'ul_flow': round(ul_flow, 1), 'dl_flow': round(dl_flow, 1),
                'ul_peak': round(ul_peak, 1), 'dl_peak': round(dl_peak, 1),
                'n_points': len(seg_df),
            })
        else:
            i += 1
    return segs

# ========== Step 2: 段合并 (复制_match_auto第988-1017行逻辑) ==========
def merge_segments(segs):
    merged = []
    for s in segs:
        if s.get('duration', 0) < 1 and s.get('flow_mb', 0) < 10:
            continue
        if not merged:
            merged.append(s)
            continue
        prev = merged[-1]
        gap = (s['start_t'] - prev['end_t']).total_seconds()
        if s['direction'] == prev['direction'] and 0 < gap < 10 and s.get('duration', 0) >= 5:
            # 合并（用原始flow值近似）
            new_dur = round((s['end_t'] - prev['start_t']).total_seconds(), 1)
            new_flow = round(s['flow_mb'] + prev['flow_mb'], 1)
            merged[-1] = {**prev,
                          'end_t': s['end_t'], 'end_i': s['end_i'],
                          'duration': new_dur, 'flow_mb': new_flow,
                          'n_points': prev['n_points'] + s['n_points']}
        else:
            merged.append(s)
    return merged

# ========== Step 3: 分类 (复制_classify_seg逻辑) ==========
def classify_seg(s):
    d = s['direction']; flow = s['flow_mb']; dur = s['duration']
    if flow < 1 and dur < 1:
        return None
    if dur <= 3:
        return 'wx_s'
    if 3 < dur <= 7:
        return 'store_s' if d == '下行' else 'wx_s'
    if 7 <= dur <= 14:
        return 'ftp_dl' if d == '下行' else 'ftp_ul'
    if 14 < dur <= 24:
        return 'store_l' if d == '下行' else 'wx_l'
    if d == '下行':
        return 'store_l'
    else:
        return 'wx_l'

# ========== Step 4: 轮次分组 (复制_match_auto第1028-1077行逻辑) ==========
def group_rounds(segs_classified):
    key_segs = []
    for s in segs_classified:
        k = s.get('key')
        if k is None: continue
        dur = s.get('duration', 0)
        if k == 'store_s':
            if dur > 25: continue
            if dur >= 5:
                s['key'] = None; continue
        if k == 'store_l' and dur < 5: continue
        if k == 'wx_s':
            key_segs.append(s)
        elif dur >= 3:
            key_segs.append(s)
        else:
            key_segs.append(s)

    rounds = []; cur_round = {}; last_end = None
    for s in key_segs:
        k = s['key']
        gap = (s['start_t'] - last_end).total_seconds() if last_end else 0
        if k in cur_round and gap <= 120:
            continue
        if len(cur_round) > 0 and (gap > 180 or (k in cur_round and gap > 120)):
            if len(cur_round) >= 2:
                rounds.append(cur_round)
            cur_round = {}
        if k not in cur_round:
            cur_round[k] = s
            last_end = s['end_t']
    if len(cur_round) >= 2:
        rounds.append(cur_round)
    return rounds

# ========== Step 5: 加载hybrid基准时间点 ==========
def load_ref_timestamps(ref_path, operator='联通'):
    wb = openpyxl.load_workbook(ref_path, data_only=True)
    ws = wb['呼叫详情']
    ref = defaultdict(list)
    def prt(raw):
        if raw is None: return None
        s = str(raw).strip()
        if not s: return None
        try: return datetime.datetime.strptime(s, '%Y-%m-%d %H:%M:%S.%f')
        except:
            try: return datetime.datetime.strptime(s, '%Y-%m-%d %H:%M:%S')
            except: return None
    def safe_float(v):
        if v is None: return 0
        try: return float(v)
        except: return 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) < 89: continue
        op = str(row[2]).strip() if row[2] else ''
        if op != operator: continue
        seq = row[7]
        ftp_s = prt(row[84]); ftp_e = prt(row[85])
        dl_r = safe_float(row[88]); ul_r = safe_float(row[87])
        if ftp_s and ftp_e and dl_r > 0:
            ref['ftp_dl'].append({'start': ftp_s.replace(microsecond=0), 'end': ftp_e.replace(microsecond=0), 'rate': dl_r, 'seq': seq})
        if ftp_s and ftp_e and ul_r > 0:
            ref['ftp_ul'].append({'start': ftp_s.replace(microsecond=0), 'end': ftp_e.replace(microsecond=0), 'rate': ul_r, 'seq': seq})
        ss = prt(row[80]); se = prt(row[81]); sr = safe_float(row[78])
        if ss and se and sr > 0:
            ref['store_s'].append({'start': ss.replace(microsecond=0), 'end': se.replace(microsecond=0), 'rate': sr, 'seq': seq})
        sl_s = prt(row[82]); sl_e = prt(row[83]); slr = safe_float(row[79])
        if sl_s and sl_e and slr > 0:
            ref['store_l'].append({'start': sl_s.replace(microsecond=0), 'end': sl_e.replace(microsecond=0), 'rate': slr, 'seq': seq})
        wl_s = prt(row[70]); wl_e = prt(row[71]); wrl = safe_float(row[66])
        if wl_s and wl_e and wrl > 0:
            ref['wx_l'].append({'start': wl_s.replace(microsecond=0), 'end': wl_e.replace(microsecond=0), 'rate': wrl, 'seq': seq})
        ws_s = prt(row[68]); ws_e = prt(row[69]); wsr = safe_float(row[65])
        if ws_s and ws_e and wsr > 0:
            ref['wx_s'].append({'start': ws_s.replace(microsecond=0), 'end': ws_e.replace(microsecond=0), 'rate': wsr, 'seq': seq})
    wb.close()
    return ref

# ========== 主诊断流程 ==========
print("=" * 80)
print("auto模式段识别各环节流失诊断")
print("=" * 80)

# 加载数据
print("\n[0] 加载MMF数据...")
df = load_mmf(MMF_PATH)
print(f"  原始行数: {len(df)}")
agg = aggregate(df)
print(f"  秒级聚合: {len(agg)} 秒")
print(f"  时间范围: {agg['t'].iloc[0]} ~ {agg['t'].iloc[-1]}")

# 加载基准
print("\n[0b] 加载hybrid基准时间点...")
ref = load_ref_timestamps(REF_PATH, '联通')
for biz in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
    print(f"  {biz}: {len(ref[biz])} 次")

# ===== 环节1: 段识别 =====
print("\n" + "=" * 80)
print("[环节1] 段识别 (gap=10, dl_min=50, ul_min=10)")
print("=" * 80)
segs_raw = detect_segments(agg, dl_min=50, ul_min=10, gap_max=10)
print(f"\n识别出 {len(segs_raw)} 个原始段")
print(f"{'#':>3} {'方向':>4} {'时长(s)':>7} {'流量MB':>8} {'UL流MB':>7} {'DL流MB':>7} {'UL峰':>6} {'DL峰':>6} {'点数':>4}  {'开始时间'}")
for i, s in enumerate(segs_raw):
    print(f"{i:>3} {s['direction']:>4} {s['duration']:>7.1f} {s['flow_mb']:>8.1f} {s['ul_flow']:>7.1f} {s['dl_flow']:>7.1f} {s['ul_peak']:>6.1f} {s['dl_peak']:>6.1f} {s['n_points']:>4}  {s['start_t']}")

# ===== 基准FTP上传 vs auto段对比 =====
print("\n" + "-" * 80)
print("[关键对比] FTP上传基准时间点 vs auto段识别结果")
print("-" * 80)
ftp_ul_refs = ref['ftp_ul']
print(f"\nFTP上传共 {len(ftp_ul_refs)} 个基准时间点:")
for ri, r in enumerate(ftp_ul_refs[:5]):
    dur_ref = (r['end'] - r['start']).total_seconds()
    # 在auto段中找最近的
    best_seg = None; best_dist = 999
    for si, s in enumerate(segs_raw):
        dist = abs((s['start_t'] - r['start']).total_seconds())
        if dist < best_dist:
            best_dist = dist; best_seg = si
    s = segs_raw[best_seg] if best_seg is not None else None
    print(f"\n  FTP上传#{ri+1}: {r['start']}~{r['end']} (dur={dur_ref:.0f}s, rate={r['rate']:.1f}Mbps)")
    if s:
        print(f"    → 最近auto段#{best_seg}: {s['start_t']}~{s['end_t']} dir={s['direction']} dur={s['duration']}s flow={s['flow_mb']}MB UL={s['ul_flow']}MB DL={s['dl_flow']}MB")
        print(f"    时间偏差: {best_dist:.0f}s")
        # 方向判定分析
        ul_f = s['ul_flow']; dl_f = s['dl_flow']
        ratio = ul_f / (dl_f + 0.01)
        correct_dir = s['direction'] == '上行'
        print(f"    UL/DL流量比: {ratio:.2f} → 方向判定: {'正确(上行)' if correct_dir else '错误(应为上行但判为下行!)'}")
        if not correct_dir:
            print(f"    *** 方向判错原因: ul_flow({ul_f}) <= dl_flow*0.5({dl_f*0.5:.1f}) ***")

# ===== 环节2: 段合并 =====
print("\n" + "=" * 80)
print("[环节2] 段合并")
print("=" * 80)
segs_merged = merge_segments(segs_raw)
print(f"\n合并后 {len(segs_merged)} 段 (原始 {len(segs_raw)} 段)")
merged_away = len(segs_raw) - len(segs_merged)
print(f"被合并/过滤: {merged_away} 段")
print(f"{'#':>3} {'方向':>4} {'时长(s)':>7} {'流量MB':>8}  {'开始时间'}")
for i, s in enumerate(segs_merged):
    print(f"{i:>3} {s['direction']:>4} {s['duration']:>7.1f} {s['flow_mb']:>8.1f}  {s['start_t']}")

# 检查FTP上传在合并后是否还在
print("\n--- FTP上传基准在合并后段的覆盖情况 ---")
ftp_ul_covered = 0
for ri, r in enumerate(ftp_ul_refs):
    found = False
    for si, s in enumerate(segs_merged):
        if s['direction'] == '上行' and abs((s['start_t'] - r['start']).total_seconds()) < 15:
            found = True
            dur_ref = (r['end'] - r['start']).total_seconds()
            print(f"  FTP上传#{ri+1}: ref={r['start']} → 合并段#{si} {s['start_t']} dur={s['duration']}s dir={s['direction']} {'OK' if found else 'MISSING'}")
            break
    if found:
        ftp_ul_covered += 1
    else:
        print(f"  FTP上传#{ri+1}: ref={r['start']} → *** 未找到匹配合并段! ***")
print(f"\nFTP上传: {ftp_ul_covered}/{len(ftp_ul_refs)} 被合并后段覆盖")

# ===== 环节3: 分类 =====
print("\n" + "=" * 80)
print("[环节3] 分类 (_classify_seg)")
print("=" * 80)
for s in segs_merged:
    s['key'] = classify_seg(s)
print(f"\n分类分布: {dict(Counter(s['key'] for s in segs_merged))}")
print(f"\n{'#':>3} {'分类':>8} {'方向':>4} {'时长':>6} {'流量MB':>8}  {'开始时间'}")
for i, s in enumerate(segs_merged):
    print(f"{i:>3} {str(s['key']):>8} {s['direction']:>4} {s['duration']:>6.1f} {s['flow_mb']:>8.1f}  {s['start_t']}")

# FTP上传分类情况
print("\n--- FTP上传的auto分类结果 ---")
ftp_ul_classified = 0
for ri, r in enumerate(ftp_ul_refs[:5]):
    best_k = None; best_dist = 999
    for si, s in enumerate(segs_merged):
        dist = abs((s['start_t'] - r['start']).total_seconds())
        if dist < best_dist:
            best_dist = dist; best_k = (si, s)
    si, s = best_k
    dur_ref = (r['end'] - r['start']).total_seconds()
    print(f"  FTP上传#{ri+1}: ref_dur={dur_ref:.0f}s → 合并段#{si} key={s['key']} dir={s['direction']} dur={s['duration']}s flow={s['flow_mb']}MB")
    if s['key'] == 'ftp_ul':
        ftp_ul_classified += 1

# ===== 环节4: 轮次分组 =====
print("\n" + "=" * 80)
print("[环节4] 轮次分组")
print("=" * 80)
rounds = group_rounds(segs_merged)
print(f"\nauto分组出 {len(rounds)} 轮")
for ri, rd in enumerate(rounds):
    keys = list(rd.keys())
    print(f"  轮#{ri+1}: {keys}")

# 每个业务在轮次中出现的次数
print(f"\n各业务在轮次中出现次数:")
for biz in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
    cnt = sum(1 for rd in rounds if biz in rd)
    bench_cnt = len(ref.get(biz, []))
    print(f"  {biz}: auto={cnt}, 基准={bench_cnt}, 流失={bench_cnt-cnt}")

# ===== 全面环节流失诊断表 =====
print("\n" + "=" * 80)
print("[全业务环节流失诊断表]")
print("=" * 80)
print(f"\n{'业务':>8} {'基准数':>6} {'环节1(段)':>10} {'环节2(合并)':>11} {'环节3(分类)':>11} {'环节4(轮次)':>11} {'流失率':>7}")

for biz, biz_name in [('ftp_dl','FTP下载'), ('ftp_ul','FTP上传'), ('store_s','商店小包'), ('store_l','商店大包'), ('wx_s','微信小包'), ('wx_l','微信大包')]:
    bench_n = len(ref.get(biz, []))
    if bench_n == 0:
        print(f"  {biz_name:>8} {bench_n:>6} {'N/A':>10} {'N/A':>11} {'N/A':>11} {'N/A':>11} {'N/A':>7}")
        continue

    # 环节1: 在原始段中有多少能匹配基准时间点(±15s, 方向一致)
    s1_hits = 0
    for r in ref[biz]:
        for s in segs_raw:
            ref_dir = '下行' if biz in ('ftp_dl', 'store_s', 'store_l') else '上行'
            if abs((s['start_t'] - r['start']).total_seconds()) < 15 and s['direction'] == ref_dir:
                s1_hits += 1
                break
    # 环节2: 合并后
    s2_hits = 0
    for r in ref[biz]:
        for s in segs_merged:
            ref_dir = '下行' if biz in ('ftp_dl', 'store_s', 'store_l') else '上行'
            if abs((s['start_t'] - r['start']).total_seconds()) < 15 and s['direction'] == ref_dir:
                s2_hits += 1
                break
    # 环节3: 分类正确
    s3_hits = 0
    for r in ref[biz]:
        for s in segs_merged:
            if abs((s['start_t'] - r['start']).total_seconds()) < 15 and s.get('key') == biz:
                s3_hits += 1
                break
    # 环节4: 进入轮次
    s4_hits = sum(1 for rd in rounds if biz in rd)

    loss_rate = (1 - s4_hits / bench_n) * 100 if bench_n > 0 else 0
    print(f"  {biz_name:>8} {bench_n:>6} {s1_hits:>10} {s2_hits:>11} {s3_hits:>11} {s4_hits:>11} {loss_rate:>6.1f}%")

# ===== 深度诊断: FTP上传每个基准时间点在auto每个环节的状态 =====
print("\n" + "=" * 80)
print("[FTP上传逐条诊断]")
print("=" * 80)
print(f"\n{'#':>3} {'基准start':>20} {'dur':>5} {'rate':>7} | {'段识别':>6} {'合并后':>6} {'分类':>8} {'问题诊断'}")
for ri, r in enumerate(ftp_ul_refs):
    dur_ref = (r['end'] - r['start']).total_seconds()
    issues = []

    # 段识别阶段
    s1_match = None
    for si, s in enumerate(segs_raw):
        if abs((s['start_t'] - r['start']).total_seconds()) < 15:
            s1_match = (si, s)
            break
    s1_str = f"#{s1_match[0]}" if s1_match else "MISS"
    if not s1_match:
        issues.append("段未识别(流量<threshold?)")
    elif s1_match[1]['direction'] != '上行':
        issues.append(f"方向判错(判为{s1_match[1]['direction']}, UL={s1_match[1]['ul_flow']}MB DL={s1_match[1]['dl_flow']}MB)")

    # 合并后
    s2_match = None
    for si, s in enumerate(segs_merged):
        if abs((s['start_t'] - r['start']).total_seconds()) < 15:
            s2_match = (si, s)
            break
    s2_str = f"#{s2_match[0]}" if s2_match else "MISS"
    if s1_match and not s2_match:
        issues.append("被合并吞掉")
    elif s1_match and s2_match and s1_match[0] != s2_match[0]:
        pass  # 段号变化是正常的

    # 分类
    s3_key = s2_match[1].get('key', '?') if s2_match else "N/A"
    if s2_match and s3_key != 'ftp_ul':
        issues.append(f"分类错(={s3_key}, dur={s2_match[1]['duration']}s)")

    print(f"{ri+1:>3} {str(r['start']):>20} {dur_ref:>5.0f} {r['rate']:>7.1f} | {s1_str:>6} {s2_str:>6} {s3_key:>8} {'; '.join(issues)}")

# ===== 方向判定深度分析 =====
print("\n" + "=" * 80)
print("[方向判定深度分析: FTP上传时间段的流量特征]")
print("=" * 80)
print("\n基准FTP上传时间段内的实际DL/UL流量:")
print(f"{'#':>3} {'start':>20} {'dur':>5} | {'DL_MAC_sum':>10} {'UL_MAC_sum':>10} {'ratio(UL/DL)':>12} {'判定':>8}")
for ri, r in enumerate(ftp_ul_refs[:10]):
    mask = (agg['t'] >= r['start']) & (agg['t'] <= r['end'])
    seg = agg[mask]
    if len(seg) == 0:
        print(f"{ri+1:>3} {str(r['start']):>20} {dur_ref:>5.0f} |  *** 时间窗口无数据 ***")
        continue
    dl_sum = float(seg['dl_mac'].sum())
    ul_sum = float(seg['ul_mac'].sum())
    ratio = ul_sum / (dl_sum + 0.01)
    auto_dir = '上行' if ul_sum > dl_sum * 0.5 else '下行'
    print(f"{ri+1:>3} {str(r['start']):>20} {(r['end']-r['start']).total_seconds():>5.0f} | {dl_sum:>10.1f} {ul_sum:>10.1f} {ratio:>12.2f} {auto_dir:>8}")

# ===== 合并吞段分析 =====
print("\n" + "=" * 80)
print("[合并吞段分析: 哪些原始段被合并了]")
print("=" * 80)
print(f"\n被合并的段（同方向+gap<10s+dur>=5s）:")
for i, s in enumerate(segs_raw):
    # 检查这个段是否被合并
    in_merged = False
    for ms in segs_merged:
        if s['start_t'] >= ms['start_t'] and s['end_t'] <= ms['end_t']:
            in_merged = True
            break
    if not in_merged:
        print(f"  原始段#{i}: {s['start_t']}~{s['end_t']} dir={s['direction']} dur={s['duration']}s flow={s['flow_mb']}MB *** 被过滤 ***")

# ===== 阈值敏感性分析 =====
print("\n" + "=" * 80)
print("[阈值敏感性分析: FTP上传识别率 vs 不同ul_min]")
print("=" * 80)
for ul_min_test in [10, 5, 3, 1, 0]:
    segs_test = detect_segments(agg, dl_min=50, ul_min=ul_min_test, gap_max=10)
    ftp_ul_hits = 0
    for r in ref['ftp_ul']:
        for s in segs_test:
            if s['direction'] == '上行' and abs((s['start_t'] - r['start']).total_seconds()) < 15:
                ftp_ul_hits += 1
                break
    print(f"  ul_min={ul_min_test:>3}: 总段数={len(segs_test):>3}, FTP上传命中={ftp_ul_hits}/{len(ref['ftp_ul'])}")

# ===== gap敏感性分析 =====
print("\n[gap敏感性分析: FTP上传识别率 vs 不同gap]")
for gap_test in [10, 5, 3, 1]:
    segs_test = detect_segments(agg, dl_min=50, ul_min=10, gap_max=gap_test)
    ftp_ul_hits = 0
    for r in ref['ftp_ul']:
        for s in segs_test:
            if s['direction'] == '上行' and abs((s['start_t'] - r['start']).total_seconds()) < 15:
                ftp_ul_hits += 1
                break
    print(f"  gap={gap_test:>2}: 总段数={len(segs_test):>3}, FTP上传命中={ftp_ul_hits}/{len(ref['ftp_ul'])}")

print("\n" + "=" * 80)
print("诊断完成")
print("=" * 80)
