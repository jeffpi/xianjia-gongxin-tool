import sys, os, traceback
import importlib.util, numpy as np, pandas as pd

spec = importlib.util.spec_from_file_location('t', '5G用户级公共监控速率统计工具-V1.04.py')
t = importlib.util.module_from_spec(spec)
spec.loader.exec_module(t)

proc = t.UCMProcessor(t.ConfigManager().config)
params = {
    'rate_column': 'RLC', 'dl_peak_limit': 900, 'ul_peak_limit': 160,
    'down_min': 50, 'up_min': 10, 'match_mode': 'hybrid',
}

proc.parse_mmf(['联通/mmf20260703115336-联通.xlsx'], params)
call_df, stats = proc.match(t.DEFAULT_PLAN, params)
round_results = proc._last_round_results

# 手动构造raw
raws = []
for p in ['联通/mmf20260703115336-联通.xlsx']:
    r = pd.read_excel(p, header=0)
    tcol = 'Time' if 'Time' in r.columns else '采集时间'
    r['_t'] = pd.to_datetime(r[tcol].astype(str).str.replace(r'\(\d+\)', '', regex=True), errors='coerce')
    r = r[r['_t'].notna()].copy()
    raws.append(r)
raw = pd.concat(raws, ignore_index=True).sort_values('_t').reset_index(drop=True)
print(f"raw: {raw.shape}", flush=True)

biz_map = [('ftp_dl', 'FTP下载'), ('ftp_ul', 'FTP上传'), ('store_s', '应用商店小包'),
           ('store_l', '应用商店大包'), ('wx_s', '微信小文件'), ('wx_l', '微信大文件')]
raw['业务标记'] = ''
biz_rows = {}
for rdata in round_results:
    for key, label in biz_map:
        if key not in rdata: continue
        s = rdata[key]; st = s['start_t']; et = s['end_t']
        if st is None or et is None: continue
        m = (raw['_t'] >= st) & (raw['_t'] <= et)
        raw.loc[m, '业务标记'] = label
        biz_rows.setdefault(label, []).extend(raw[m].index.tolist())

biz_rate_info = {
    'FTP下载': ('下载平均速率(Mbps)', '下行', '业务时长(秒)'),
    'FTP上传': ('上传平均速率(Mbps)', '上行', '业务时长(秒)'),
    '应用商店小包': ('小文件下载速率(Mbps)', '下行', '小文件下载时长(秒)'),
    '应用商店大包': ('大文件下载速率(Mbps)', '下行', '大文件下载时长(秒)'),
    '微信小文件': ('微信小包发送速率(Mbps)', '上行', '微信小包时长(秒)'),
    '微信大文件': ('微信大包发送速率(Mbps)', '上行', '微信大包时长(秒)'),
}
biz_rate_vals = {}
for rdata in round_results:
    for key, label in biz_map:
        if key not in rdata: continue
        s = rdata[key]; st = s['start_t']; et = s['end_t']
        if st is None or et is None: continue
        dur = round((et - st).total_seconds(), 3)
        rate = s.get('rate', 0)
        if rate and dur:
            biz_rate_vals[label] = (rate, dur)

# 检查重叠
labels = list(biz_rows.keys())
overlaps_found = False
for i in range(len(labels)):
    for j in range(i+1, len(labels)):
        s1 = set(biz_rows[labels[i]])
        s2 = set(biz_rows[labels[j]])
        inter = s1 & s2
        if inter:
            overlaps_found = True
            print(f'重叠: {labels[i]} x {labels[j]}: {len(inter)}行', flush=True)

if not overlaps_found:
    print('无重叠', flush=True)

# 检查非连续
for label in labels:
    idxs = sorted(set(biz_rows[label]))
    gaps = sum(1 for i in range(len(idxs)-1) if idxs[i+1] - idxs[i] > 1)
    print(f'{label}: {len(idxs)}行, 间隙={gaps}', flush=True)

print('DONE', flush=True)