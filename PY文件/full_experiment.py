#!/usr/bin/env python3
"""
全口径实验 - 精确按呼叫详情切割，输出结果文件对标
"""
import sys,os,numpy as np,pandas as pd
sys.stderr=open(os.devnull,'w')

f_zb='/Users/sun/ClaudeCode/先甲工信部工具/电信/工信部业务指标V2120260428_202607030329.xlsx'
cd=pd.read_excel(f_zb,sheet_name='呼叫详情')

# 输出结果文件基准值  [rate_rc, rate_clip, ...]
other={'联通':{'FTP下载':{'rate':711.114,'clip':610.377},'FTP上传':{'rate':88.131,'clip':87.732},
              '应用商店_小包':{'rate':143.738,'dur':4.364},'应用商店_大包':{'rate':511.491,'dur':15.7},
              '微信_小文件':{'rate':22.676,'dur':1.6},'微信_大文件':{'rate':44.127,'dur':30}},
       '电信':{'FTP下载':{'rate':855.944,'clip':717.793},'FTP上传':{'rate':124.493,'clip':123.357},
              '应用商店_小包':{'rate':135.16,'dur':4.077},'应用商店_大包':{'rate':497.263,'dur':13.727},
              '微信_小文件':{'rate':23.493,'dur':1.538},'微信_大文件':{'rate':43.844,'dur':29.643}}}

# 定义6种业务和对应的呼叫详情类型
biz_defs={
    'FTP下载':('FTPDownload','下行','dl_rlc','dl_clip','dl_mac'),
    'FTP上传':('FTPUpload','上行','ul_rlc','ul_clip','ul_mac'),
    '应用商店_小包':('StoreDownloadSmall','下行','dl_rlc','dl_clip','dl_mac'),
    '应用商店_大包':('StoreDownloadLarge','下行','dl_rlc','dl_clip','dl_mac'),
    '微信_小文件':('WeChatSmall','上行','ul_rlc','ul_clip','ul_mac'),
    '微信_大文件':('WeChatLarge','上行','ul_rlc','ul_clip','ul_mac'),
}

# 口径策略
strategies=[
    ('rlc_mean',  lambda seg,rc,cc,mc: float(seg[rc].mean())),
    ('rlc_nz',    lambda seg,rc,cc,mc: float(seg[rc].replace(0,np.nan).dropna().mean()) if len(seg[rc].replace(0,np.nan).dropna()) else 0),
    ('clip_mean', lambda seg,rc,cc,mc: float(seg[cc].mean())),
    ('clip_nz',   lambda seg,rc,cc,mc: float(seg[cc].replace(0,np.nan).dropna().mean()) if len(seg[cc].replace(0,np.nan).dropna()) else 0),
    ('mac_mean',  lambda seg,rc,cc,mc: float(seg[mc].mean())),
    ('mac_nz',    lambda seg,rc,cc,mc: float(seg[mc].replace(0,np.nan).dropna().mean()) if len(seg[mc].replace(0,np.nan).dropna()) else 0),
    ('rlc_max',   lambda seg,rc,cc,mc: float(seg[rc].max())),
    ('clip_max',  lambda seg,rc,cc,mc: float(seg[cc].max())),
    ('mac_max',   lambda seg,rc,cc,mc: float(seg[mc].max())),
    ('rlc_all_nz_max', lambda seg,rc,cc,mc: float(np.mean([v for v in seg[rc] if v>0])) if len(seg[rc][seg[rc]>0]) else 0),
    ('clip_all_nz_max', lambda seg,rc,cc,mc: float(np.mean([v for v in seg[cc] if v>0])) if len(seg[cc][seg[cc]>0]) else 0),
]

# 预加载UCM数据
import importlib.util
spec=importlib.util.spec_from_file_location('t','/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py')
t=importlib.util.module_from_spec(spec); spec.loader.exec_module(t)

for label,files in [('联通',['联通/mmf20260703115336-联通.xlsx']),
                     ('电信',['电信/mmf20260703115334-电信.xlsx','电信/mmf20260703115328-电信.xlsx'])]:
    proc=t.UCMProcessor(t.ConfigManager().config)
    params={'rate_column':'RLC','dl_peak_limit':900,'ul_peak_limit':160,'down_min':50,'up_min':10,'match_mode':'auto'}
    proc.parse_mmf(files,params)
    agg=proc.seconds

    print(f'\n{"="*90}')
    print(f'{label}')
    print(f'{"业务":<14} {"ref_rate":>10} ', end='')
    for sn,fn in strategies:
        print(f'{sn:>12}', end='')
    print()

    for biz_name,(biz_type,dir_key,rlc_c,clip_c,mac_c) in biz_defs.items():
        ref=other[label].get(biz_name,{}).get('rate',0)
        if ref==0: continue

        # 按呼叫详情精准切割并计算各口径
        all_rates={sn:[] for sn,fn in strategies}
        n=0
        for i in range(1,len(cd)):
            row=cd.iloc[i]; op=row.iloc[2]; btype=row.iloc[4]
            if not (op==label and btype==biz_type): continue
            start=pd.Timestamp(row.iloc[84]); end=pd.Timestamp(row.iloc[85])
            mask=(agg['t']>=start)&(agg['t']<end); seg=agg[mask]
            if len(seg)<2: continue
            n+=1
            for sn,fn in strategies:
                all_rates[sn].append(fn(seg,rlc_c,clip_c,mac_c))

        if n==0:
            print(f'{biz_name:<14} {ref:>10.1f} {"N/A":>12}')
            continue

        print(f'{biz_name:<14} {ref:>10.1f} ', end='')
        best_dev=999; best_name=''
        for sn,fn in strategies:
            vals=all_rates[sn]; avg=float(np.mean(vals)); dev=(avg-ref)/ref*100 if ref else 0
            mk='✓' if abs(dev)<10 else '±' if abs(dev)<20 else '✗'
            print(f'{avg:>8.1f}%{dev:+.0f}m ', end='')
            if abs(dev)<abs(best_dev): best_dev=dev; best_name=sn
        print(f'\n  BEST={best_name} dev={best_dev:+.1f}%')
PYEOF