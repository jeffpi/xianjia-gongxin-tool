#!/usr/bin/env python3
"""V2.00算法独立验证"""
import sys,os,numpy as np,pandas as pd,datetime,openpyxl

sys.stderr=open(os.devnull,'w')

def parse_time(raw):
    if raw is None: return None
    s=str(raw).strip()
    if not s: return None
    try:
        clean=__import__('re').sub(r'\(\d+\)','',s).strip()
        return datetime.datetime.strptime(clean,'%Y-%m-%d %H:%M:%S')
    except: return None

def safe_int(raw):
    if raw is None: return 0
    s=str(raw).strip()
    if not s or s.upper()=='N/A': return 0
    try: return int(float(s))
    except: return 0

def safe_mbps(raw):
    if raw is None: return 0.0
    s=str(raw).strip()
    if not s or s.upper()=='N/A': return 0.0
    try: return float(s)/1_000_000
    except: return 0.0

def load_mmf(paths):
    all_rows=[]
    for fp in paths:
        wb=openpyxl.load_workbook(fp,data_only=True)
        ws=wb.worksheets[0]
        for row in ws.iter_rows(min_row=8,values_only=True):
            if len(row)<107: continue
            dt=parse_time(row[2])
            if dt is None: continue
            crnti=safe_int(row[9])
            if crnti==0: continue
            all_rows.append({'time':dt,'crnti':crnti,
                'dl_mac':safe_mbps(row[101]),'ul_mac':safe_mbps(row[102]),
                'dl_rlc':safe_mbps(row[105]),'ul_rlc':safe_mbps(row[106])})
        wb.close()
    groups={}
    for r in all_rows:
        key=(r['time'],r['crnti'])
        if key not in groups:
            groups[key]={'dl_mac':0.0,'ul_mac':0.0,'dl_rlc':0.0,'ul_rlc':0.0}
        g=groups[key]; g['dl_mac']+=r['dl_mac']; g['ul_mac']+=r['ul_mac']
        g['dl_rlc']=max(g['dl_rlc'],r['dl_rlc']); g['ul_rlc']=max(g['ul_rlc'],r['ul_rlc'])
    result=[]
    for (t,c),v in groups.items():
        result.append({'time':t,'crnti':c,'dl_mac':v['dl_mac'],'ul_mac':v['ul_mac'],
                       'dl_rlc':v['dl_rlc'],'ul_rlc':v['ul_rlc']})
    result.sort(key=lambda r:r['time'])
    return result

def prt(raw):
    if raw is None: return None
    s=str(raw).strip()
    if not s: return None
    try: return datetime.datetime.strptime(s,'%Y-%m-%d %H:%M:%S.%f')
    except:
        try: return datetime.datetime.strptime(s,'%Y-%m-%d %H:%M:%S')
        except: return None

base='/Users/sun/ClaudeCode/先甲工信部工具'

# 加载参考文件
ref_path=os.path.join(base,'联通','工信部业务指标V2120260428_202607030329.xlsx')
wb=openpyxl.load_workbook(ref_path,data_only=True)
ws=wb['呼叫详情']
ref={'电信':{},'联通':{}}
for op in ref:
    ref[op]={'ftp_dl':[],'ftp_ul':[],'store_s':[],'store_l':[],'wx_s':[],'wx_l':[]}

for row in ws.iter_rows(min_row=3,values_only=True):
    if len(row)<89: continue
    op=str(row[2]).strip() if row[2] else ''
    if op not in ('电信','联通'): continue
    ftp_s=prt(row[84]); ftp_e=prt(row[85])
    dl_r=float(row[88]) if row[88] and str(row[88]).strip().replace('.','').replace('-','').isdigit() else 0
    ul_r=float(row[87]) if row[87] and str(row[87]).strip().replace('.','').replace('-','').isdigit() else 0
    if ftp_s and ftp_e and dl_r>0: ref[op]['ftp_dl'].append({'start':ftp_s,'end':ftp_e,'rate':dl_r})
    if ftp_s and ftp_e and ul_r>0: ref[op]['ftp_ul'].append({'start':ftp_s,'end':ftp_e,'rate':ul_r})
    ss=prt(row[80]); se=prt(row[81]); sr=float(row[78]) if row[78] and str(row[78]).strip().replace('.','').replace('-','').isdigit() else 0
    if ss and se and sr>0: ref[op]['store_s'].append({'start':ss,'end':se,'rate':sr})
    ls=prt(row[82]); le=prt(row[83]); lr=float(row[79]) if row[79] and str(row[79]).strip().replace('.','').replace('-','').isdigit() else 0
    if ls and le and lr>0: ref[op]['store_l'].append({'start':ls,'end':le,'rate':lr})
    # 微信文件 col62=微信文件业务请求时间(61) ... col67=微信大包发送速率(66)，col68=结果(67)
    # col69=小包开始(68)，col70=小包完成(69)，col71=大包开始(70)，col72=大包完成(71)
    wxs=prt(row[68]); wxe=prt(row[69])
    wrs=0
    if row[65] is not None:
        try: wrs=float(row[65])
        except: pass
    if wxs and wxe and wrs>0: ref[op]['wx_s'].append({'start':wxs,'end':wxe,'rate':wrs})
    wxl=prt(row[70]); wxl_e=prt(row[71])
    wrl=0
    if row[66] is not None:
        try: wrl=float(row[66])
        except: pass
    if wxl and wxl_e and wrl>0: ref[op]['wx_l'].append({'start':wxl,'end':wxl_e,'rate':wrl})
wb.close()

BIZ_CONFIG=[
    ('ftp_dl','FTP下载','dl','dl_mac',1000),
    ('ftp_ul','FTP上传','ul','ul_mac',200),
    ('store_s','应用商店_小包','dl','dl_rlc',1000),
    ('store_l','应用商店_大包','dl','dl_mac_nz',None),
    ('wx_s','微信_小文件','ul','ul_mac_nz',None),
    ('wx_l','微信_大包','ul','ul_rlc',200),
]

# 输出结果文件参考值
REF_OUT={
    '联通':{'FTP下载':711.114,'FTP上传':88.131,'应用商店_小包':143.738,'应用商店_大包':511.491,'微信_小文件':22.676,'微信_大包':44.127},
    '电信':{'FTP下载':855.944,'FTP上传':124.493,'应用商店_小包':135.16,'应用商店_大包':497.263,'微信_小文件':23.493,'微信_大包':43.844},
}

for label, files in [('联通',[os.path.join(base,'联通','mmf20260703115336-联通.xlsx')]),
                     ('电信',[os.path.join(base,'电信','mmf20260703115334-电信.xlsx'),
                             os.path.join(base,'电信','mmf20260703115328-电信.xlsx')])]:
    mmf=load_mmf(files)
    print(f'\n{"="*60}')
    print(f'{label} V2.00 (参考时间戳+动态CRNTI)')
    print(f'{"="*60}')
    print(f'{"业务":<12} {"N":>3} {"计算均值":>10} {"参考均值":>10} {"vs参考":>7} {"输出文件":>10} {"vs输出":>7}')
    print('-'*60)

    for biz_key,name,dir_col,rate_col,clip in BIZ_CONFIG:
        ref_rows=ref[label][biz_key]
        results=[]
        for r in ref_rows:
            w=[x for x in mmf if r['start'].replace(microsecond=0)<=x['time']<r['end'].replace(microsecond=0)]
            if not w: continue
            wdf=pd.DataFrame(w)
            traffic=wdf.groupby('crnti')[dir_col+'_mac'].sum()
            if len(traffic)==0: continue
            best=traffic.idxmax()
            cdf=wdf[wdf['crnti']==best]
            if '_nz' in rate_col:
                vals=cdf[rate_col.replace('_nz','')].replace(0,np.nan).dropna()
            else:
                vals=cdf[rate_col].clip(upper=clip) if clip else cdf[rate_col]
            if len(vals)==0: continue
            results.append({'calc':vals.mean(),'ref':r['rate']})

        if not results:
            print(f'{name:<12} {"0":>3} {"N/A":>10} {"N/A":>10} {"N/A":>7} {"N/A":>10}')
            continue

        n=len(results)
        cm=np.mean([x['calc'] for x in results])
        rm=np.mean([x['ref'] for x in results])
        dev=(cm-rm)/rm*100 if rm else 0
        of=REF_OUT[label].get(name,0)
        dev_of=(cm-of)/of*100 if of else 0
        mk='✓' if abs(dev)<10 else '±' if abs(dev)<15 else '✗'
        mk2='✓' if abs(dev_of)<10 else '±' if abs(dev_of)<15 else '✗'
        # V2.00与参考文件指标的对齐程度
        print(f'{name:<12} {n:>3} {cm:>10.1f} {rm:>10.1f} {dev:>+6.1f}%{mk} {of:>10.1f} {dev_of:>+6.1f}%{mk2}')