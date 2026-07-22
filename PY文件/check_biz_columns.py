#!/usr/bin/env python3
"""检查呼叫详情中商店/微信业务的列索引"""
import pandas as pd

f_zb='/Users/sun/ClaudeCode/先甲工信部工具/电信/工信部业务指标V2120260428_202607030329.xlsx'
cd=pd.read_excel(f_zb,sheet_name='呼叫详情')

# 打印表头
print('列数:', len(cd.columns))
print('表头:', list(cd.columns))

# 检查几个StoreDownloadSmall的行
cnt=0
for i in range(1,min(50,len(cd))):
    if cnt>=3: break
    row=cd.iloc[i]
    op=str(row.iloc[2]); biz=str(row.iloc[4])
    if 'Store' in biz or 'WeChat' in biz:
        # 时间已经在后面几列
        t1=row.iloc[84]; t2=row.iloc[85]
        dl=row.iloc[88]; ul=row.iloc[87]
        print(f'row{i}: op={op} biz={biz} start={t1} end={t2} dl={dl} ul={ul}')
PYEOF