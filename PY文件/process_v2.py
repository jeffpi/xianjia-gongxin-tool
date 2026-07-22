#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V2处理脚本：对齐基准 + 计算FTP识别规则
"""

import pandas as pd
import re
from datetime import datetime, timedelta

# 读取文件
df = pd.read_excel('联通/mmf20260703-电信-合并-带基准.xlsx')
baseline = pd.read_excel('工信部业务指标_呼叫详情_整理.xlsx', sheet_name='电信')
ftp_baseline = baseline[baseline['业务类型'].str.contains('FTP')]

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

# 解析C列时间
df['_c_time'] = df['Time'].apply(parse_time)

# 初始化列（使用object类型避免dtype问题）
g_cols = ['基准_业务类型', '基准_开始时间', '基准_结束时间', '基准_速率类指标', '基准_数值', '基准_时长类指标', '基准_数值.1']
for col in g_cols:
    df[col] = pd.Series([None] * len(df), dtype='object')
df['业务识别'] = pd.Series([None] * len(df), dtype='object')
df['持续时长'] = pd.Series([None] * len(df), dtype='object')

print(f"总行数: {len(df)}")
print(f"FTP基准业务数: {len(ftp_baseline)}")

# 对齐基准数据
matched = 0
unmatched = []
aligned_segments = []  # 保存成功对齐的段

for _, row in ftp_baseline.iterrows():
    start_time = pd.to_datetime(row['开始时间'])
    end_time = pd.to_datetime(row['结束时间'])

    # 找最接近开始时间的C列行
    best_start = None
    min_diff = 9999.0
    for i, c_time in enumerate(df['_c_time']):
        if c_time is not None:
            diff = abs((c_time - start_time).total_seconds())
            if diff < min_diff:
                min_diff = diff
                best_start = i

    # 找最接近结束时间的C列行
    best_end = None
    min_diff_end = 9999.0
    for i, c_time in enumerate(df['_c_time']):
        if c_time is not None and i >= best_start:
            diff = abs((c_time - end_time).total_seconds())
            if diff < min_diff_end:
                min_diff_end = diff
                best_end = i

    # 只有开始误差<5秒才算成功对齐
    if best_start is not None and best_end is not None and min_diff < 5:
        for idx in range(best_start, min(best_end + 1, len(df))):
            df.at[idx, '基准_业务类型'] = str(row['业务类型'])
            df.at[idx, '基准_开始时间'] = str(row['开始时间'])
            df.at[idx, '基准_结束时间'] = str(row['结束时间'])
            df.at[idx, '基准_速率类指标'] = str(row['速率类指标'])
            df.at[idx, '基准_数值'] = str(row['数值'])
            df.at[idx, '基准_时长类指标'] = str(row['时长类指标'])
            df.at[idx, '基准_数值.1'] = str(row['数值.1'])

        c_dur = (df.loc[best_end, '_c_time'] - df.loc[best_start, '_c_time']).total_seconds()
        aligned_segments.append({
            'biz': str(row['业务类型']),
            'start_row': best_start,
            'end_row': best_end,
            'c_dur': c_dur,
            'b_dur': (end_time - start_time).total_seconds()
        })
        matched += 1
    else:
        unmatched.append(f"{row['业务类型']}: {start_time.strftime('%H:%M:%S')} (误差{min_diff:.1f}s)")

print(f"\n成功对齐: {matched} 个FTP段")
print(f"未对齐: {len(unmatched)} 个")
for u in unmatched:
    print(f"  - {u}")

# 分析C列时长（只统计成功对齐的）
durations = [s['c_dur'] for s in aligned_segments]
print(f"\nC列时长范围（已对齐）: {min(durations):.3f}s ~ {max(durations):.3f}s")
print(f"C列时长均值: {sum(durations)/len(durations):.3f}s")

# 显示对齐的段
print(f"\n对齐的FTP段:")
for i, seg in enumerate(aligned_segments[:20], 1):
    print(f"  {i}. {seg['biz']}: 行{seg['start_row']}-{seg['end_row']} C列时长{seg['c_dur']:.3f}s 基准时长{seg['b_dur']:.3f}s")

# ============================================
# 用基准的开始/结束时间来直接标记F列
# 当前阶段：建立算法过程中，可以用基准时间
# ============================================
print(f"\n用基准时间标记F列（业务识别）...")
for seg in aligned_segments:
    for idx in range(seg['start_row'], seg['end_row'] + 1):
        df.at[idx, '业务识别'] = seg['biz']
        df.at[idx, '持续时长'] = round(seg['c_dur'], 3)

# 显示前20个段
print(f"\nF列前20个FTP段:")
prev = ''
cnt = 0
for _, row in df.iterrows():
    if row['业务识别'] and row['业务识别'] != prev:
        cnt += 1
        if cnt <= 20:
            print(f"  {cnt}. {row['业务识别']}: {str(row['Time'])[:19]} 持续{row['持续时长']}s")
        prev = row['业务识别']

# 重新排列列
cols = df.columns.tolist()
time_idx = cols.index('Time')
new_cols = cols[:time_idx+1] + ['下载速率', '上传速率', '持续时长', '业务识别'] + g_cols + [c for c in cols if c not in set(cols[:time_idx+1] + ['下载速率', '上传速率', '持续时长', '业务识别'] + g_cols + ['_c_time'])]
df = df[new_cols]

output_file = '联通/mmf20260703-电信-合并-带基准_V2.xlsx'
df.to_excel(output_file, index=False)
print(f"\n已保存: {output_file}")

ftp_biz = df[df['业务识别'].notna() & (df['业务识别'] != '')]
print(f"\nFTP总行数: {len(ftp_biz)}")
print(f"FTP下载: {(ftp_biz['业务识别']=='FTP下载').sum()}")
print(f"FTP上传: {(ftp_biz['业务识别']=='FTP上传').sum()}")
