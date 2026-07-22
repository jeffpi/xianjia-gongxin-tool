"""调试多业务标记重叠"""
import sys, os, traceback
sys.stderr = open(os.devnull, 'w')
import importlib.util, pandas as pd

spec = importlib.util.spec_from_file_location('t', '5G用户级公共监控速率统计工具-V1.04.py')
t = importlib.util.module_from_spec(spec)
spec.loader.exec_module(t)

proc = t.UCMProcessor(t.ConfigManager().config)
params = {
    'rate_column': 'RLC', 'dl_peak_limit': 900, 'ul_peak_limit': 160,
    'down_min': 50, 'up_min': 10, 'match_mode': 'hybrid',
}
proc.parse_mmf(['联通/mmf20260703115336-联通.xlsx'], params)
_, stats = proc.match(t.DEFAULT_PLAN, params)
rr = proc._last_round_results

# 原始数据
raws = []
r = pd.read_excel('联通/mmf20260703115336-联通.xlsx', header=0)
tcol = '采集时间'
r['_t'] = pd.to_datetime(r[tcol].astype(str).str.replace(r'\(\d+\)', '', regex=True), errors='coerce')
r = r[r['_t'].notna()].copy()
raws.append(r)
raw = pd.concat(raws, ignore_index=True).sort_values('_t').reset_index(drop=True)

biz_map = [('ftp_dl', 'FTP下载'), ('ftp_ul', 'FTP上传'), ('store_s', '应用商店小包'),
           ('store_l', '应用商店大包'), ('wx_s', '微信小文件'), ('wx_l', '微信大文件')]

all_idxs = {}
for rdata in rr:
    for key, label in biz_map:
        if key not in rdata:
            continue
        s = rdata[key]
        st = s['start_t']
        et = s['end_t']
        if st is None or et is None:
            continue
        m = (raw['_t'] >= st) & (raw['_t'] <= et)
        for idx in raw[m].index:
            all_idxs.setdefault(idx, []).append(label)

multi = [(idx, labels) for idx, labels in all_idxs.items() if len(set(labels)) > 1]
print(f'多业务标记行: {len(multi)}', flush=True)
for idx, labels in multi[:15]:
    print(f'  索引{idx}: {set(labels)} t={raw.iloc[idx]["_t"]}', flush=True)