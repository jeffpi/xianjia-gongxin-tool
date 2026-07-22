#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V3处理脚本：标记所有业务 + 轮次列
"""

import pandas as pd
import re
from datetime import datetime, timedelta

# 读取文件
df = pd.read_excel('联通/mmf20260703-电信-合并-带基准.xlsx')
baseline = pd.read_excel('工信部业务指标_呼叫详情_整理.xlsx', sheet_name='电信')

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

# 初始化列
g_cols = ['基准_业务类型', '基准_开始时间', '基准_结束时间', '基准_速率类指标', '基准_数值', '基准_时长类指标', '基准_数值.1']
for col in g_cols:
    df[col] = pd.Series([None] * len(df), dtype='object')
df['业务识别'] = pd.Series([None] * len(df), dtype='object')
df['持续时长'] = pd.Series([None] * len(df), dtype='object')
df['轮次'] = pd.Series([None] * len(df), dtype='object')

# 按时间排序基准文件
baseline_sorted = baseline.sort_values('开始时间').reset_index(drop=True)
print(f"基准业务总数: {len(baseline_sorted)}")
print(f"总轮数: {len(baseline_sorted) // 6}")

# 计算每轮：每6个业务为1轮
total_rounds = len(baseline_sorted) // 6

matched = 0
unmatched = []

for round_idx in range(total_rounds):
    round_num = round_idx + 1
    round_data = baseline_sorted.iloc[round_idx * 6:(round_idx + 1) * 6]

    print(f"\n轮次 {round_num}:")
    for _, row in round_data.iterrows():
        biz_type = str(row['业务类型'])
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

        if best_start is not None and best_end is not None and min_diff < 5:
            c_dur = (df.loc[best_end, '_c_time'] - df.loc[best_start, '_c_time']).total_seconds()

            for idx in range(best_start, min(best_end + 1, len(df))):
                df.at[idx, '基准_业务类型'] = biz_type
                df.at[idx, '基准_开始时间'] = str(row['开始时间'])
                df.at[idx, '基准_结束时间'] = str(row['结束时间'])
                df.at[idx, '基准_速率类指标'] = str(row['速率类指标'])
                df.at[idx, '基准_数值'] = str(row['数值'])
                df.at[idx, '基准_时长类指标'] = str(row['时长类指标'])
                df.at[idx, '基准_数值.1'] = str(row['数值.1'])
                df.at[idx, '业务识别'] = biz_type
                df.at[idx, '持续时长'] = round(c_dur, 3)
                df.at[idx, '轮次'] = round_num

            print(f"  {biz_type}: 行{best_start}-{best_end} 时长{c_dur:.3f}s")
            matched += 1
        else:
            unmatched.append(f"轮{round_num} {biz_type}: {start_time.strftime('%H:%M:%S')} (误差{min_diff:.1f}s)")

print(f"\n成功对齐: {matched} 个业务段")
print(f"未对齐: {len(unmatched)} 个")
for u in unmatched:
    print(f"  - {u}")

# 重新排列列
# A-C: 原有
# D: 下载速率
# E: 上传速率
# 持续时长
# F: 业务识别
# G: 基准_业务类型
# H: 基准_开始时间
# 轮次（插入H和I之间）
# I: 基准_结束时间
# J-M: 其它基准列

cols = df.columns.tolist()
time_idx = cols.index('Time')

# 构建新列顺序
new_cols = cols[:time_idx+1]  # A-C
new_cols += ['下载速率', '上传速率', '持续时长', '业务识别']  # D-F
new_cols += ['基准_业务类型', '基准_开始时间']  # G-H
new_cols += ['轮次']  # 新列（H和I之间）
new_cols += ['基准_结束时间', '基准_速率类指标', '基准_数值', '基准_时长类指标', '基准_数值.1']  # I-M

# 其它列
used = set(new_cols + ['_c_time'])
other_cols = [c for c in cols if c not in used]
new_cols += other_cols

df = df[new_cols]

# 保存
output_file = '联通/mmf20260703-电信-合并-带基准_V3.xlsx'
df.to_excel(output_file, index=False)
print(f"\n已保存: {output_file}")

# 统计
biz_rows = df[df['业务识别'].notna() & (df['业务识别'] != '')]
print(f"\n业务标记总行数: {len(biz_rows)}")
print(f"\n各业务行数:")
print(biz_rows['业务识别'].value_counts())
print(f"\n各轮次行数:")
print(biz_rows['轮次'].value_counts().sort_index())

# 显示轮次概览
print(f"\n各轮次业务概览:")
for r in range(1, min(total_rounds + 1, 6)):
    round_data = biz_rows[biz_rows['轮次'] == r]
    biz_list = round_data['业务识别'].unique()
    print(f"  轮次{r}: {list(biz_list)}")
