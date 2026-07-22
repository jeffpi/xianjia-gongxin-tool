#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V2.0.0 业务识别测试 - 简化版
从MMF数据自动识别业务，不依赖基准文件
"""

import os
import re
import pandas as pd
from datetime import datetime, timedelta
from collections import Counter

# ==================== 配置 ====================
# 业务识别阈值
DL_RATE_THRESHOLD = 10.0  # Mbps，下行高速阈值
UL_RATE_THRESHOLD = 10.0  # Mbps，上行高速阈值
MIN_FTP_DURATION = 5.0     # 秒，FTP最小时长
MAX_FTP_DURATION = 15.0    # 秒，FTP最大时长

# ==================== 时间解析 ====================
def parse_time(time_str):
    """解析MMF时间格式：2026-07-03 02:02:54(099)"""
    try:
        time_str = str(time_str).strip()
        match = re.match(r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\((\d+)\)', time_str)
        if match:
            dt = datetime.strptime(match.group(1), '%Y-%m-%d %H:%M:%S')
            ms = int(match.group(2))
            return dt + timedelta(milliseconds=ms)
        return pd.NaT
    except:
        return pd.NaT

# ==================== 加载数据 ====================
def load_mmf(file_path):
    """加载MMF文件"""
    print(f"\n加载: {os.path.basename(file_path)}")
    
    df = pd.read_excel(file_path)
    
    # 解析时间
    df['timestamp'] = df['Time'].apply(parse_time)
    df = df[df['timestamp'].notna()].copy()
    
    # 提取速率列
    dl_col = 'Downlink MAC Throughput(bps)'
    ul_col = 'Uplink MAC Throughput(bps)'
    
    if dl_col in df.columns:
        df['dl_mac'] = pd.to_numeric(df[dl_col], errors='coerce') / 1e6
    if ul_col in df.columns:
        df['ul_mac'] = pd.to_numeric(df[ul_col], errors='coerce') / 1e6
    
    print(f"  数据行数: {len(df)}")
    print(f"  时间范围: {df['timestamp'].min()} ~ {df['timestamp'].max()}")
    
    return df

# ==================== 业务识别 ====================
def identify_business_segments(df):
    """识别业务片段"""
    print("\n识别业务片段...")
    
    df = df.sort_values('timestamp').reset_index(drop=True)
    
    segments = []
    
    # 识别下行高速段
    dl_segments = find_high_speed_segments(df, 'dl_mac', DL_RATE_THRESHOLD)
    print(f"  下行高速段: {len(dl_segments)} 个")
    
    # 识别上行高速段  
    ul_segments = find_high_speed_segments(df, 'ul_mac', UL_RATE_THRESHOLD, min_duration=0.5)
    print(f"  上行高速段: {len(ul_segments)} 个")
    
    # 分类下行段
    for seg in dl_segments:
        biz = classify_download_segment(df, seg)
        if biz:
            segments.append(biz)
    
    # 分类上行段
    for seg in ul_segments:
        biz = classify_upload_segment(df, seg)
        if biz:
            segments.append(biz)
    
    # 按时间排序
    segments.sort(key=lambda x: x['start_time'])
    
    print(f"  识别业务: {len(segments)} 个")
    
    return segments

def find_high_speed_segments(df, rate_col, threshold, min_duration=1.0):
    """查找高速段"""
    segments = []
    
    high_speed = df[rate_col] > threshold
    
    in_segment = False
    start_idx = None
    start_time = None
    
    for i, is_high in enumerate(high_speed):
        if is_high and not in_segment:
            in_segment = True
            start_idx = i
            start_time = df.iloc[i]['timestamp']
        elif not is_high and in_segment:
            in_segment = False
            end_time = df.iloc[i]['timestamp']
            duration = (end_time - start_time).total_seconds()
            
            if duration >= min_duration:
                segments.append({
                    'start_idx': start_idx,
                    'end_idx': i - 1,
                    'start_time': start_time,
                    'end_time': end_time,
                    'duration_sec': duration,
                })
    
    # 处理最后一个段
    if in_segment:
        end_time = df.iloc[-1]['timestamp']
        duration = (end_time - start_time).total_seconds()
        if duration >= min_duration:
            segments.append({
                'start_idx': start_idx,
                'end_idx': len(df) - 1,
                'start_time': start_time,
                'end_time': end_time,
                'duration_sec': duration,
            })
    
    return segments

def classify_download_segment(df, segment):
    """分类下行段"""
    start_idx = segment['start_idx']
    end_idx = segment['end_idx']
    duration = segment['duration_sec']
    
    seg_df = df.iloc[start_idx:end_idx+1]
    rates = seg_df['dl_mac'].dropna()
    rates = rates[rates > 0]
    
    if len(rates) == 0:
        return None
    
    avg_rate = rates.mean()
    
    # 分类规则
    if MIN_FTP_DURATION <= duration <= MAX_FTP_DURATION:
        biz_type = "FTP下载"
    elif duration < 5:
        biz_type = "商店小文件下载"
    else:
        biz_type = "商店大文件下载"
    
    return {
        'type': biz_type,
        'start_time': segment['start_time'],
        'end_time': segment['end_time'],
        'duration_sec': duration,
        'avg_rate': round(avg_rate, 2),
    }

def classify_upload_segment(df, segment):
    """分类上行段"""
    start_idx = segment['start_idx']
    end_idx = segment['end_idx']
    duration = segment['duration_sec']
    
    seg_df = df.iloc[start_idx:end_idx+1]
    rates = seg_df['ul_mac'].dropna()
    rates = rates[rates > 0]
    
    if len(rates) == 0:
        return None
    
    avg_rate = rates.mean()
    
    # 分类规则
    if MIN_FTP_DURATION <= duration <= MAX_FTP_DURATION:
        biz_type = "FTP上传"
    elif duration < 5:
        biz_type = "微信小包发送"
    else:
        biz_type = "微信大包发送"
    
    return {
        'type': biz_type,
        'start_time': segment['start_time'],
        'end_time': segment['end_time'],
        'duration_sec': duration,
        'avg_rate': round(avg_rate, 2),
    }

# ==================== 与基准对比 ====================
def compare_with_baseline(segments, baseline_file, operator='电信'):
    """与基准文件对比"""
    print(f"\n与基准对比...")
    
    df_baseline = pd.read_excel(baseline_file, sheet_name=operator)
    print(f"  基准业务数: {len(df_baseline)}")
    
    results = []
    
    for seg in segments:
        start_time = seg['start_time']
        end_time = seg['end_time']
        biz_type = seg['type']
        calc_rate = seg['avg_rate']
        calc_duration = seg['duration_sec']
        
        # 查找最接近的基准业务
        best_match = None
        min_time_diff = float('inf')
        
        for _, row in df_baseline.iterrows():
            baseline_start = pd.to_datetime(row['开始时间'])
            time_diff = abs((start_time - baseline_start).total_seconds())
            
            if time_diff < min_time_diff and time_diff < 30:  # 30秒容差
                min_time_diff = time_diff
                best_match = row
        
        if best_match is not None:
            baseline_rate = best_match['数值']
            baseline_duration = best_match['数值.1']
            
            rate_deviation = (calc_rate - baseline_rate) / baseline_rate * 100 if baseline_rate > 0 else None
            duration_deviation = (calc_duration - baseline_duration) / baseline_duration * 100 if baseline_duration > 0 else None
            
            results.append({
                '业务类型': biz_type,
                '开始时间': start_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                '结束时间': end_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                '计算时长(秒)': round(calc_duration, 2),
                '基准时长(秒)': round(baseline_duration, 2),
                '时长偏差%': round(duration_deviation, 2) if duration_deviation else None,
                '计算速率(Mbps)': calc_rate,
                '基准速率(Mbps)': baseline_rate,
                '速率偏差%': round(rate_deviation, 2) if rate_deviation else None,
                '状态': '达标' if rate_deviation and abs(rate_deviation) <= 10 else '偏差',
            })
        else:
            results.append({
                '业务类型': biz_type,
                '开始时间': start_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                '结束时间': end_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3],
                '计算时长(秒)': round(calc_duration, 2),
                '基准时长(秒)': None,
                '时长偏差%': None,
                '计算速率(Mbps)': calc_rate,
                '基准速率(Mbps)': None,
                '速率偏差%': None,
                '状态': '未匹配',
            })
    
    return pd.DataFrame(results)

# ==================== 主程序 ====================
def main():
    print("=" * 80)
    print("V2.0.0 业务识别测试 - 简化版")
    print("=" * 80)
    
    # 输入文件
    file1 = "/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115328-电信.xlsx"
    file2 = "/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115334-电信.xlsx"
    baseline_file = "/Users/sun/ClaudeCode/先甲工信部工具/工信部业务指标_呼叫详情_整理.xlsx"
    
    # 加载数据
    df1 = load_mmf(file1)
    df2 = load_mmf(file2)
    
    # 合并
    df = pd.concat([df1, df2], ignore_index=True)
    df = df.sort_values('timestamp').reset_index(drop=True)
    print(f"\n合并后: {len(df)} 行")
    
    # 业务识别
    segments = identify_business_segments(df)
    
    # 统计业务类型
    print("\n业务类型分布:")
    biz_types = Counter([s['type'] for s in segments])
    for name, count in sorted(biz_types.items()):
        print(f"  {name}: {count} 个")
    
    # 与基准对比
    df_compare = compare_with_baseline(segments, baseline_file, '电信')
    
    # 统计对比结果
    print("\n对比统计:")
    print(f"  总识别业务: {len(df_compare)}")
    print(f"  达标: {len(df_compare[df_compare['状态'] == '达标'])}")
    print(f"  偏差: {len(df_compare[df_compare['状态'] == '偏差'])}")
    print(f"  未匹配: {len(df_compare[df_compare['状态'] == '未匹配'])}")
    
    # 保存结果
    output_file = "/Users/sun/ClaudeCode/先甲工信部工具/V2自动识别对比结果.xlsx"
    df_compare.to_excel(output_file, index=False)
    print(f"\n结果已保存: {output_file}")
    
    # 显示前10个结果
    print("\n前10个识别结果:")
    for i, row in df_compare.head(10).iterrows():
        print(f"  {i+1}. {row['业务类型']}: {row['开始时间'][11:19]} ~ {row['结束时间'][11:19]} | 速率偏差: {row['速率偏差%']}% | {row['状态']}")

if __name__ == "__main__":
    main()
