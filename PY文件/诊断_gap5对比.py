#!/usr/bin/env python3
"""补充诊断: gap=5 vs gap=10 全流程对比 + 方向判定根因深挖"""
import sys, os, re, datetime
import pandas as pd
import numpy as np
import openpyxl
from collections import Counter, defaultdict

TOOL_DIR = "/Users/sun/ClaudeCode/先甲工信部工具"
MMF_PATH = os.path.join(TOOL_DIR, "联通/mmf20260703115336-联通.xlsx")
REF_PATH = os.path.join(TOOL_DIR, "联通/工信部业务指标V2120260428_202607030329.xlsx")

def parse_t(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ('', 'nan', 'None'):
        return None
    try: return pd.to_datetime(re.sub(r'\(\d+\)', '', str(v)).strip())
    except: return None

def select_col(df, direction, layer):
    dir_kw = ['下行', 'Downlink'] if direction == '下行' else ['上行', 'Uplink']
    cands = [c for c in df.columns if any(k in str(c) for k in dir_kw)
             and layer in str(c) and ('吞吐率' in str(c) or 'hroughput' in str(c).lower())]
    return cands[0] if cands else None

def find_col(df, keywords):
    for c in df.columns:
        if any(k in str(c) for k in keywords): return c
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

def aggregate(df):
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
    agg['dl_clip'] = agg['dl_rlc'].clip(upper=900)
    agg['ul_clip'] = agg['ul_rlc'].clip(upper=160)
    agg[['dl_mac', 'ul_mac', 'dl_rlc', 'ul_rlc']] = agg[['dl_mac', 'ul_mac', 'dl_rlc', 'ul_rlc']].fillna(0)
    return agg

def detect_segments(agg, dl_min=50, ul_min=10, gap_max=10):
    n = len(agg); segs = []; i = 0
    while i < n:
        if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
            start_i = i; gap = 0
            while i < n and gap < gap_max:
                if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min: gap = 0
                else: gap += 1
                i += 1
            end_i = max(start_i, i - 1 - gap)
            seg_df = agg.loc[start_i:end_i]
            seg_dur = (seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds()
            ul_flow = float(seg_df['ul_mac'].sum()) / 8
            dl_flow = float(seg_df['dl_mac'].sum()) / 8
            ul_peak = float(seg_df['ul_mac'].max())
            dl_peak = float(seg_df['dl_mac'].max())
            if ul_flow > dl_flow * 0.5 or (seg_dur <= 3 and ul_peak > 15 and dl_peak < ul_peak):
                direction = '上行'; col = 'ul_mac'
            else:
                direction = '下行'; col = 'dl_mac'
            flow = seg_df[col].sum() / 8
            dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
            segs.append({'start_i': start_i, 'end_i': end_i,
                         'start_t': seg_df['t'].iloc[0], 'end_t': seg_df['t'].iloc[-1],
                         'direction': direction, 'flow_mb': round(flow, 1), 'duration': round(dur, 1),
                         'ul_flow': round(ul_flow, 1), 'dl_flow': round(dl_flow, 1),
                         'ul_peak': round(ul_peak, 1), 'dl_peak': round(dl_peak, 1), 'n_points': len(seg_df)})
        else: i += 1
    return segs

def merge_segments(segs):
    merged = []
    for s in segs:
        if s.get('duration', 0) < 1 and s.get('flow_mb', 0) < 10: continue
        if not merged: merged.append(s); continue
        prev = merged[-1]
        gap = (s['start_t'] - prev['end_t']).total_seconds()
        if s['direction'] == prev['direction'] and 0 < gap < 10 and s.get('duration', 0) >= 5:
            merged[-1] = {**prev, 'end_t': s['end_t'], 'end_i': s['end_i'],
                          'duration': round((s['end_t'] - prev['start_t']).total_seconds(), 1),
                          'flow_mb': round(s['flow_mb'] + prev['flow_mb'], 1),
                          'n_points': prev['n_points'] + s['n_points']}
        else: merged.append(s)
    return merged

def classify_seg(s):
    d = s['direction']; flow = s['flow_mb']; dur = s['duration']
    if flow < 1 and dur < 1: return None
    if dur <= 3: return 'wx_s'
    if 3 < dur <= 7: return 'store_s' if d == '下行' else 'wx_s'
    if 7 <= dur <= 14: return 'ftp_dl' if d == '下行' else 'ftp_ul'
    if 14 < dur <= 24: return 'store_l' if d == '下行' else 'wx_l'
    return 'store_l' if d == '下行' else 'wx_l'

def group_rounds(segs_classified):
    key_segs = []
    for s in segs_classified:
        k = s.get('key')
        if k is None: continue
        dur = s.get('duration', 0)
        if k == 'store_s':
            if dur > 25: continue
            if dur >= 5: s['key'] = None; continue
        if k == 'store_l' and dur < 5: continue
        key_segs.append(s)
    rounds = []; cur_round = {}; last_end = None
    for s in key_segs:
        k = s['key']
        gap = (s['start_t'] - last_end).total_seconds() if last_end else 0
        if k in cur_round and gap <= 120: continue
        if len(cur_round) > 0 and (gap > 180 or (k in cur_round and gap > 120)):
            if len(cur_round) >= 2: rounds.append(cur_round)
            cur_round = {}
        if k not in cur_round:
            cur_round[k] = s; last_end = s['end_t']
    if len(cur_round) >= 2: rounds.append(cur_round)
    return rounds

def load_ref(ref_path, operator='联通'):
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
        try: return float(v) if v else 0
        except: return 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) < 89: continue
        op = str(row[2]).strip() if row[2] else ''
        if op != operator: continue
        ftp_s = prt(row[84]); ftp_e = prt(row[85])
        dl_r = safe_float(row[88]); ul_r = safe_float(row[87])
        if ftp_s and ftp_e and dl_r > 0: ref['ftp_dl'].append({'start': ftp_s.replace(microsecond=0), 'end': ftp_e.replace(microsecond=0), 'rate': dl_r})
        if ftp_s and ftp_e and ul_r > 0: ref['ftp_ul'].append({'start': ftp_s.replace(microsecond=0), 'end': ftp_e.replace(microsecond=0), 'rate': ul_r})
        ss = prt(row[80]); se = prt(row[81]); sr = safe_float(row[78])
        if ss and se and sr > 0: ref['store_s'].append({'start': ss.replace(microsecond=0), 'end': se.replace(microsecond=0), 'rate': sr})
        sl_s = prt(row[82]); sl_e = prt(row[83]); slr = safe_float(row[79])
        if sl_s and sl_e and slr > 0: ref['store_l'].append({'start': sl_s.replace(microsecond=0), 'end': sl_e.replace(microsecond=0), 'rate': slr})
        wl_s = prt(row[70]); wl_e = prt(row[71]); wrl = safe_float(row[66])
        if wl_s and wl_e and wrl > 0: ref['wx_l'].append({'start': wl_s.replace(microsecond=0), 'end': wl_e.replace(microsecond=0), 'rate': wrl})
        ws_s = prt(row[68]); ws_e = prt(row[69]); wsr = safe_float(row[65])
        if ws_s and ws_e and wsr > 0: ref['wx_s'].append({'start': ws_s.replace(microsecond=0), 'end': ws_e.replace(microsecond=0), 'rate': wsr})
    wb.close()
    return ref

# ===== 加载 =====
print("加载数据...")
df = load_mmf(MMF_PATH)
agg = aggregate(df)
ref = load_ref(REF_PATH, '联通')

BIZ_DIR = {'ftp_dl':'下行','ftp_ul':'上行','store_s':'下行','store_l':'下行','wx_s':'上行','wx_l':'上行'}
BIZ_NAME = {'ftp_dl':'FTP下载','ftp_ul':'FTP上传','store_s':'商店小包','store_l':'商店大包','wx_s':'微信小包','wx_l':'微信大包'}

# ===== 全流程对比: gap=10 vs gap=5 =====
print("\n" + "=" * 80)
print("[全流程对比: gap=10 vs gap=5]")
print("=" * 80)

for gap_val in [10, 5]:
    segs_raw = detect_segments(agg, dl_min=50, ul_min=10, gap_max=gap_val)
    segs_merged = merge_segments(segs_raw)
    for s in segs_merged: s['key'] = classify_seg(s)
    rounds = group_rounds(segs_merged)

    print(f"\n--- gap={gap_val} ---")
    print(f"原始段: {len(segs_raw)}, 合并后: {len(segs_merged)}, 轮次: {len(rounds)}")
    cls_dist = Counter(s['key'] for s in segs_merged)
    print(f"分类分布: {dict(cls_dist)}")

    print(f"\n{'业务':>8} {'基准':>5} {'环节1':>6} {'环节2':>6} {'环节3':>6} {'环节4':>6} {'流失率':>7}")
    for biz in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
        bench_n = len(ref.get(biz, []))
        ref_dir = BIZ_DIR[biz]
        # 环节1
        s1 = sum(1 for r in ref[biz] for s in segs_raw
                 if abs((s['start_t']-r['start']).total_seconds())<15 and s['direction']==ref_dir)
        # 环节2
        s2 = sum(1 for r in ref[biz] for s in segs_merged
                 if abs((s['start_t']-r['start']).total_seconds())<15 and s['direction']==ref_dir)
        # 环节3
        s3 = sum(1 for r in ref[biz] for s in segs_merged
                 if abs((s['start_t']-r['start']).total_seconds())<15 and s.get('key')==biz)
        # 环节4
        s4 = sum(1 for rd in rounds if biz in rd)
        loss = (1-s4/bench_n)*100 if bench_n else 0
        print(f"  {BIZ_NAME[biz]:>8} {bench_n:>5} {s1:>6} {s2:>6} {s3:>6} {s4:>6} {loss:>6.1f}%")

# ===== gap=5时FTP上传逐条分析 =====
print("\n" + "=" * 80)
print("[gap=5: FTP上传逐条分析]")
print("=" * 80)
segs5 = detect_segments(agg, dl_min=50, ul_min=10, gap_max=5)
print(f"\nFTP上传基准 {len(ref['ftp_ul'])} 条, gap=5时识别 {len(segs5)} 段")
print(f"\n{'#':>3} {'FTP_UL基准start':>20} {'dur':>5} {'rate':>7} | {'匹配段#':>6} {'方向':>4} {'段dur':>6} {'UL流':>7} {'DL流':>7} {'分类':>8} {'问题'}")
for ri, r in enumerate(ref['ftp_ul']):
    best = None; best_dist = 999
    for si, s in enumerate(segs5):
        d = abs((s['start_t']-r['start']).total_seconds())
        if d < best_dist: best_dist = d; best = (si, s)
    issues = []
    if not best or best_dist > 15:
        print(f"{ri+1:>3} {str(r['start']):>20} {(r['end']-r['start']).total_seconds():>5.0f} {r['rate']:>7.1f} |  *** 未匹配(dist={best_dist:.0f}s) ***")
        continue
    si, s = best
    if s['direction'] != '上行':
        issues.append(f"方向错(UL={s['ul_flow']}/DL={s['dl_flow']})")
    # 分类
    cls = classify_seg(s)
    if cls != 'ftp_ul':
        issues.append(f"分类错(={cls},dur={s['duration']}s)")
    print(f"{ri+1:>3} {str(r['start']):>20} {(r['end']-r['start']).total_seconds():>5.0f} {r['rate']:>7.1f} | #{si:>5} {s['direction']:>4} {s['duration']:>6.1f} {s['ul_flow']:>7.1f} {s['dl_flow']:>7.1f} {cls:>8} {'; '.join(issues)}")

# ===== 关键: 每轮业务的时间结构分析 =====
print("\n" + "=" * 80)
print("[业务时间结构: 以第一轮为例, 逐秒DL/UL流量]")
print("=" * 80)
# 第一轮FTP下载开始时间
first_ftp = ref['ftp_dl'][0]['start']
# 展示前后120秒的逐秒流量
mask = (agg['t'] >= first_ftp - pd.Timedelta(seconds=5)) & (agg['t'] <= first_ftp + pd.Timedelta(seconds=120))
seg = agg[mask].reset_index(drop=True)
print(f"\n第一轮开始: {first_ftp}")
print(f"{'时间':>20} {'DL_MAC':>8} {'UL_MAC':>8} {'备注':>20}")
for _, row in seg.iterrows():
    t = row['t']
    dl = row['dl_mac']; ul = row['ul_mac']
    note = ''
    for biz in ref:
        for r in ref[biz]:
            if r['start'] <= t <= r['end']:
                note = f"{biz} ({r['rate']:.0f}Mbps)"
                break
        if note: break
    print(f"{str(t):>20} {dl:>8.1f} {ul:>8.1f} {note:>20}")

# ===== 方向判定: ul_flow > dl_flow*0.5 阈值分析 =====
print("\n" + "=" * 80)
print("[方向判定阈值分析: FTP上传的UL/DL比值分布]")
print("=" * 80)
print(f"\nFTP上传窗口内UL_flow/DL_flow比值分布:")
ratios = []
for ri, r in enumerate(ref['ftp_ul']):
    mask = (agg['t'] >= r['start']) & (agg['t'] <= r['end'])
    seg = agg[mask]
    if len(seg) == 0: continue
    dl_sum = float(seg['dl_mac'].sum())
    ul_sum = float(seg['ul_mac'].sum())
    ratio = ul_sum / (dl_sum + 0.01)
    ratios.append(ratio)
print(f"  样本数: {len(ratios)}")
print(f"  最小比值: {min(ratios):.2f}")
print(f"  最大比值: {max(ratios):.2f}")
print(f"  中位数: {sorted(ratios)[len(ratios)//2]:.2f}")
print(f"  >0.5的: {sum(1 for r in ratios if r > 0.5)}/{len(ratios)}")
print(f"  >0.3的: {sum(1 for r in ratios if r > 0.3)}/{len(ratios)}")
print(f"  >0.2的: {sum(1 for r in ratios if r > 0.2)}/{len(ratios)}")
print(f"  >0.1的: {sum(1 for r in ratios if r > 0.1)}/{len(ratios)}")

# ===== gap=5 + 方向阈值修正后的模拟 =====
print("\n" + "=" * 80)
print("[模拟修正: gap=5 + 方向阈值=0.15]")
print("=" * 80)

def detect_segments_v2(agg, dl_min=50, ul_min=10, gap_max=5, dir_ratio=0.15):
    """改进版: gap=5, 方向阈值=0.15"""
    n = len(agg); segs = []; i = 0
    while i < n:
        if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min:
            start_i = i; gap = 0
            while i < n and gap < gap_max:
                if agg.at[i, 'dl_mac'] > dl_min or agg.at[i, 'ul_mac'] > ul_min: gap = 0
                else: gap += 1
                i += 1
            end_i = max(start_i, i - 1 - gap)
            seg_df = agg.loc[start_i:end_i]
            seg_dur = (seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds()
            ul_flow = float(seg_df['ul_mac'].sum()) / 8
            dl_flow = float(seg_df['dl_mac'].sum()) / 8
            ul_peak = float(seg_df['ul_mac'].max())
            dl_peak = float(seg_df['dl_mac'].max())
            # 方向判定：ratio降低到0.15
            if ul_flow > dl_flow * dir_ratio or (seg_dur <= 3 and ul_peak > 15 and dl_peak < ul_peak):
                direction = '上行'; col = 'ul_mac'
            else:
                direction = '下行'; col = 'dl_mac'
            flow = seg_df[col].sum() / 8
            dur = max(round((seg_df['t'].iloc[-1] - seg_df['t'].iloc[0]).total_seconds(), 1), 0.5)
            segs.append({'start_i': start_i, 'end_i': end_i,
                         'start_t': seg_df['t'].iloc[0], 'end_t': seg_df['t'].iloc[-1],
                         'direction': direction, 'flow_mb': round(flow, 1), 'duration': round(dur, 1),
                         'ul_flow': round(ul_flow, 1), 'dl_flow': round(dl_flow, 1),
                         'ul_peak': round(ul_peak, 1), 'dl_peak': round(dl_peak, 1), 'n_points': len(seg_df)})
        else: i += 1
    return segs

for gap_val, dir_ratio in [(5, 0.15), (5, 0.10), (5, 0.20), (3, 0.15)]:
    segs_raw = detect_segments_v2(agg, dl_min=50, ul_min=10, gap_max=gap_val, dir_ratio=dir_ratio)
    segs_merged = merge_segments(segs_raw)
    for s in segs_merged: s['key'] = classify_seg(s)
    rounds = group_rounds(segs_merged)
    cls_dist = Counter(s['key'] for s in segs_merged)

    print(f"\n--- gap={gap_val}, dir_ratio={dir_ratio} ---")
    print(f"段: {len(segs_raw)}→{len(segs_merged)}, 分类: {dict(cls_dist)}, 轮次: {len(rounds)}")
    for biz in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
        bench_n = len(ref.get(biz, []))
        s4 = sum(1 for rd in rounds if biz in rd)
        print(f"  {BIZ_NAME[biz]:>8}: {s4}/{bench_n} ({s4/bench_n*100:.0f}%)")

# ===== store_s的dur>=5被过滤的影响 =====
print("\n" + "=" * 80)
print("[store_s dur过滤影响分析]")
print("=" * 80)
segs5 = detect_segments(agg, dl_min=50, ul_min=10, gap_max=5)
segs5m = merge_segments(segs5)
for s in segs5m: s['key'] = classify_seg(s)
store_s_segs = [s for s in segs5m if s.get('key') == 'store_s']
print(f"\n分类为store_s的段: {len(store_s_segs)} 个")
print(f"{'#':>3} {'dur':>5} {'flow':>7} {'开始时间':>20} {'过滤后?'}")
for i, s in enumerate(store_s_segs):
    filtered = '过滤' if s['duration'] >= 5 else '保留'
    print(f"{i:>3} {s['duration']:>5.1f} {s['flow_mb']:>7.1f} {str(s['start_t']):>20} {filtered}")

# 商店小包基准时长分析
print(f"\n商店小包基准时长:")
durs = [(r['end']-r['start']).total_seconds() for r in ref['store_s']]
print(f"  样本: {len(durs)}, 范围: {min(durs):.0f}~{max(durs):.0f}s, 中位: {sorted(durs)[len(durs)//2]:.0f}s")
print(f"  >=5s的: {sum(1 for d in durs if d>=5)}/{len(durs)}")

# ===== wx_s过滤分析 =====
print("\n" + "=" * 80)
print("[wx_s 过滤分析]")
print("=" * 80)
wx_s_raw = [s for s in segs5 if s['direction']=='上行' and s['duration']<=3]
print(f"\ngap=5时上行短段(dur<=3): {len(wx_s_raw)}")
print(f"{'#':>3} {'dur':>5} {'flow':>7} {'过滤后?'}")
for i, s in enumerate(wx_s_raw[:20]):
    filtered = '过滤(dur<1&flow<10)' if s['duration']<1 and s['flow_mb']<10 else '保留'
    print(f"{i:>3} {s['duration']:>5.1f} {s['flow_mb']:>7.1f} {filtered}")
filtered_cnt = sum(1 for s in wx_s_raw if s['duration']<1 and s['flow_mb']<10)
print(f"\n被过滤: {filtered_cnt}/{len(wx_s_raw)}")

# wx_s基准
durs_wxs = [(r['end']-r['start']).total_seconds() for r in ref['wx_s']]
print(f"\n微信小包基准时长: 范围={min(durs_wxs):.0f}~{max(durs_wxs):.0f}s, 中位={sorted(durs_wxs)[len(durs_wxs)//2]:.0f}s")

print("\n诊断完成")
