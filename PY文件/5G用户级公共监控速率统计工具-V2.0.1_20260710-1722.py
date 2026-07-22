#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
5G用户级公共监控速率统计工具 V2.0.1
简化版 - 可直接运行
"""

import os
import sys
import pandas as pd
import re
from datetime import datetime, timedelta

# ==================== 时间解析 ====================
def parse_time_str(time_str):
    """解析时间：2026-07-03 02:02:54(099)"""
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

def load_mmf_file(file_path):
    """加载MMF文件"""
    print(f"加载: {os.path.basename(file_path)}")
    
    df = pd.read_excel(file_path, header=0)
    
    # 解析时间
    df['timestamp'] = [parse_time_str(t) for t in df['Time']]
    
    # 提取速率列
    dl_col = 'Downlink MAC Throughput(bps)'
    ul_col = 'Uplink MAC Throughput(bps)'
    
    if dl_col in df.columns:
        df['dl_mac'] = df[dl_col] / 1e6  # bps -> Mbps
    if ul_col in df.columns:
        df['ul_mac'] = df[ul_col] / 1e6
    
    # 过滤无效时间
    df = df[df['timestamp'].notna()]
    
    print(f"  数据行数: {len(df)}")
    print(f"  时间范围: {df['timestamp'].min()} ~ {df['timestamp'].max()}")
    
    return df

# ==================== 业务识别 ====================
def identify_businesses(df, dl_threshold=10.0, ul_threshold=10.0):
    """识别业务"""
    df = df.sort_values('timestamp').reset_index(drop=True)
    
    businesses = []
    
    # 识别下行高速段
    dl_high = df['dl_mac'] > dl_threshold
    dl_segments = find_segments(df, dl_high, '下行')
    
    # 识别上行高速段
    ul_high = df['ul_mac'] > ul_threshold
    ul_segments = find_segments(df, ul_high, '上行')
    
    # 分类下行段
    for seg in dl_segments:
        seg_df = df.iloc[seg['start']:seg['end']+1]
        avg_rate = seg_df['dl_mac'].mean()
        duration = seg['duration']
        
        # 分类
        if 5 <= duration <= 15:
            biz_type = "FTP下载"
        elif duration < 5:
            biz_type = "商店小文件下载"
        else:
            biz_type = "商店大文件下载"
        
        businesses.append({
            'type': biz_type,
            'start_time': seg['start_time'],
            'end_time': seg['end_time'],
            'duration_sec': duration,
            'avg_rate': round(avg_rate, 2),
        })
    
    # 分类上行段
    for seg in ul_segments:
        seg_df = df.iloc[seg['start']:seg['end']+1]
        avg_rate = seg_df['ul_mac'].mean()
        duration = seg['duration']
        
        if 5 <= duration <= 15:
            biz_type = "FTP上传"
        elif duration < 5:
            biz_type = "微信小包发送"
        else:
            biz_type = "微信大包发送"
        
        businesses.append({
            'type': biz_type,
            'start_time': seg['start_time'],
            'end_time': seg['end_time'],
            'duration_sec': duration,
            'avg_rate': round(avg_rate, 2),
        })
    
    # 按时间排序
    businesses.sort(key=lambda x: x['start_time'])
    
    return businesses

def find_segments(df, condition, direction, min_duration=1.0):
    """查找连续段"""
    segments = []
    in_seg = False
    start_idx = None
    start_time = None
    
    for i, is_high in enumerate(condition):
        if is_high and not in_seg:
            in_seg = True
            start_idx = i
            start_time = df.iloc[i]['timestamp']
        elif not is_high and in_seg:
            in_seg = False
            end_time = df.iloc[i]['timestamp']
            duration = (end_time - start_time).total_seconds()
            
            if duration >= min_duration:
                segments.append({
                    'start': start_idx,
                    'end': i - 1,
                    'start_time': start_time,
                    'end_time': end_time,
                    'duration': duration,
                    'direction': direction,
                })
    
    # 处理最后一个段
    if in_seg:
        end_time = df.iloc[-1]['timestamp']
        duration = (end_time - start_time).total_seconds()
        if duration >= min_duration:
            segments.append({
                'start': start_idx,
                'end': len(df) - 1,
                'start_time': start_time,
                'end_time': end_time,
                'duration': duration,
                'direction': direction,
            })
    
    return segments

# ==================== 主程序 ====================
def main():
    print("=" * 60)
    print("V2.0.1 业务识别测试")
    print("=" * 60)
    
    # 输入文件
    file1 = "/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115328-电信.xlsx"
    file2 = "/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115334-电信.xlsx"
    
    # 加载
    df1 = load_mmf_file(file1)
    df2 = load_mmf_file(file2)
    
    # 合并
    df = pd.concat([df1, df2], ignore_index=True)
    df = df.sort_values('timestamp').reset_index(drop=True)
    print(f"\n合并后: {len(df)} 行")
    
    # 识别
    print("\n识别业务...")
    businesses = identify_businesses(df)
    
    print(f"\n识别结果: {len(businesses)} 个业务")
    
    # 统计
    biz_types = Counter([b['type'] for b in businesses])
    print("\n业务类型分布:")
    for name, count in sorted(biz_types.items()):
        print(f"  {name}: {count} 个")
    
    # 显示前10个
    print("\n前10个业务:")
    for i, biz in enumerate(businesses[:10]):
        print(f"  {i+1}. {biz['type']}: {biz['start_time'].strftime('%H:%M:%S')} ~ {biz['end_time'].strftime('%H:%M:%S')} ({biz['duration_sec']:.1f}s, {biz['avg_rate']:.1f} Mbps)")
    
    # 保存结果
    output_file = "/Users/sun/ClaudeCode/先甲工信部工具/V2识别结果.xlsx"
    df_result = pd.DataFrame(businesses)
    df_result.to_excel(output_file, index=False)
    print(f"\n结果已保存: {output_file}")

if __name__ == "__main__":
    main()
