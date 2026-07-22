#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
再生脚本：从输入文件的呼叫详情sheet重新生成输出文件
修复问题：各业务时长现为硬编码值，改为从实际开始/结束时间计算
"""

import pandas as pd
from datetime import datetime, timedelta
import os

# 输入输出路径
INPUT_FILE = '联通/汇总-工信部业务指标V2120260428_202607170548.xlsx'
OUTPUT_FILE = '联通/工信部业务指标_呼叫详情_整理-20260718.xlsx'

# 列名映射
BIZ_MAP = {
    'FTP下载': {
        'test_biz': 'FTPDownload',
        'start_col': '发起时间',
        'end_col': '完成时间',
        'rate_col': '下载平均速率(Mbps)',
        'rate_label': '下载平均速率(Mbps)',
        'dur_label': '业务时长(秒)',
    },
    'FTP上传': {
        'test_biz': 'FTPUpload',
        'start_col': '发起时间',
        'end_col': '完成时间',
        'rate_col': '上传平均速率(Mbps)',
        'rate_label': '上传平均速率(Mbps)',
        'dur_label': '业务时长(秒)',
    },
    '微信小包发送': {
        'test_biz': None,  # 无测试业务标记，靠微信小包发送速率非空
        'start_col': '微信小包发送开始时间',
        'end_col': '微信小包发送完成时间',
        'rate_col': '微信小包发送速率(Mbps)',
        'rate_label': '微信小包发送速率(Mbps)',
        'dur_label': '业务时长(秒)',
    },
    '微信大包发送': {
        'test_biz': None,
        'start_col': '微信大包发送开始时间',
        'end_col': '微信大包发送完成时间',
        'rate_col': '微信大包发送速率(Mbps)',
        'rate_label': '微信大包发送速率(Mbps)',
        'dur_label': '业务时长(秒)',
    },
    '应用商店小文件下载': {
        'test_biz': None,
        'start_col': '应用商店小文件下载开始时间',
        'end_col': '应用商店小文件下载完成时间',
        'rate_col': '应用商店小文件下载速率(Mbps)',
        'rate_label': '应用商店小文件下载速率(Mbps)',
        'dur_label': '业务时长(秒)',
    },
    '应用商店大文件下载': {
        'test_biz': None,
        'start_col': '应用商店大文件下载开始时间',
        'end_col': '应用商店大文件下载完成时间',
        'rate_col': '应用商店大文件下载速率(Mbps)',
        'rate_label': '应用商店大文件下载速率(Mbps)',
        'dur_label': '业务时长(秒)',
    },
}

def parse_time(t):
    """解析时间列，支持多种格式"""
    if pd.isna(t) or t is None or t == '' or t == 'nan':
        return None
    try:
        t = str(t).strip()
        # 尝试解析带毫秒的时间
        for fmt in ['%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S']:
            try:
                return datetime.strptime(t, fmt)
            except ValueError:
                continue
        # 尝试直接pd.to_datetime
        return pd.to_datetime(t)
    except:
        return None

def compute_duration(start, end):
    """计算时长(秒)"""
    if start is None or end is None:
        return None
    diff = (end - start).total_seconds()
    return round(diff, 3)

def main():
    print("=" * 60)
    print("再生脚本：从呼叫详情重新生成输出文件")
    print("=" * 60)

    # 读取输入文件
    print(f"\n读取输入文件: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE, sheet_name='呼叫详情', header=None)

    # 第1行是列名
    col_names = df.iloc[1].tolist()
    data = df.iloc[2:].copy()
    data.columns = col_names
    data = data.reset_index(drop=True)
    print(f"  呼叫详情行数: {len(data)}")

    # 存储提取的所有行
    rows = []

    # 1. FTP下载 - 按测试业务筛选
    print("\n[1/6] 提取FTP下载...")
    ftp_dl = data[data['测试业务'] == 'FTPDownload'].copy()
    ftp_dl = ftp_dl[ftp_dl['发起时间'].notna()]  # 排除发起时间为空的行
    cnt = 0
    for _, row in ftp_dl.iterrows():
        start = parse_time(row['发起时间'])
        end = parse_time(row['完成时间'])
        if start is None or end is None:
            continue
        dur = compute_duration(start, end)
        rows.append({
            '运营商': row['运营商'],
            '业务类型': 'FTP下载',
            '开始时间': start,
            '结束时间': end,
            '速率类指标': '下载平均速率(Mbps)',
            '数值': round(float(row['下载平均速率(Mbps)']), 3) if pd.notna(row['下载平均速率(Mbps)']) else 0,
            '时长类指标': '业务时长(秒)',
            '数值.1': dur,
        })
        cnt += 1
    print(f"  提取 {cnt} 行")

    # 2. FTP上传 - 按测试业务筛选
    print("[2/6] 提取FTP上传...")
    ftp_ul = data[data['测试业务'] == 'FTPUpload'].copy()
    ftp_ul = ftp_ul[ftp_ul['发起时间'].notna()]
    cnt = 0
    for _, row in ftp_ul.iterrows():
        start = parse_time(row['发起时间'])
        end = parse_time(row['完成时间'])
        if start is None or end is None:
            continue
        dur = compute_duration(start, end)
        rows.append({
            '运营商': row['运营商'],
            '业务类型': 'FTP上传',
            '开始时间': start,
            '结束时间': end,
            '速率类指标': '上传平均速率(Mbps)',
            '数值': round(float(row['上传平均速率(Mbps)']), 3) if pd.notna(row['上传平均速率(Mbps)']) else 0,
            '时长类指标': '业务时长(秒)',
            '数值.1': dur,
        })
        cnt += 1
    print(f"  提取 {cnt} 行")

    # 3. 微信小包发送 - 按微信小包发送速率非空筛选
    print("[3/6] 提取微信小包发送...")
    wx_small = data[data['微信小包发送速率(Mbps)'].notna()].copy()
    cnt = 0
    for _, row in wx_small.iterrows():
        start = parse_time(row['微信小包发送开始时间'])
        end = parse_time(row['微信小包发送完成时间'])
        if start is None or end is None:
            continue
        dur = compute_duration(start, end)
        rows.append({
            '运营商': row['运营商'],
            '业务类型': '微信小包发送',
            '开始时间': start,
            '结束时间': end,
            '速率类指标': '微信小包发送速率(Mbps)',
            '数值': round(float(row['微信小包发送速率(Mbps)']), 3) if pd.notna(row['微信小包发送速率(Mbps)']) else 0,
            '时长类指标': '业务时长(秒)',
            '数值.1': dur,
        })
        cnt += 1
    print(f"  提取 {cnt} 行")

    # 4. 微信大包发送 - 按微信大包发送速率非空筛选
    print("[4/6] 提取微信大包发送...")
    wx_large = data[data['微信大包发送速率(Mbps)'].notna()].copy()
    cnt = 0
    for _, row in wx_large.iterrows():
        start = parse_time(row['微信大包发送开始时间'])
        end = parse_time(row['微信大包发送完成时间'])
        if start is None or end is None:
            continue
        dur = compute_duration(start, end)
        rows.append({
            '运营商': row['运营商'],
            '业务类型': '微信大包发送',
            '开始时间': start,
            '结束时间': end,
            '速率类指标': '微信大包发送速率(Mbps)',
            '数值': round(float(row['微信大包发送速率(Mbps)']), 3) if pd.notna(row['微信大包发送速率(Mbps)']) else 0,
            '时长类指标': '业务时长(秒)',
            '数值.1': dur,
        })
        cnt += 1
    print(f"  提取 {cnt} 行")

    # 5. 应用商店小文件下载
    print("[5/6] 提取应用商店小文件下载...")
    app_small = data[data['应用商店小文件下载速率(Mbps)'].notna()].copy()
    cnt = 0
    for _, row in app_small.iterrows():
        start = parse_time(row['应用商店小文件下载开始时间'])
        end = parse_time(row['应用商店小文件下载完成时间'])
        if start is None or end is None:
            continue
        dur = compute_duration(start, end)
        rows.append({
            '运营商': row['运营商'],
            '业务类型': '应用商店小文件下载',
            '开始时间': start,
            '结束时间': end,
            '速率类指标': '应用商店小文件下载速率(Mbps)',
            '数值': round(float(row['应用商店小文件下载速率(Mbps)']), 3) if pd.notna(row['应用商店小文件下载速率(Mbps)']) else 0,
            '时长类指标': '业务时长(秒)',
            '数值.1': dur,
        })
        cnt += 1
    print(f"  提取 {cnt} 行")

    # 6. 应用商店大文件下载
    print("[6/6] 提取应用商店大文件下载...")
    app_large = data[data['应用商店大文件下载速率(Mbps)'].notna()].copy()
    cnt = 0
    for _, row in app_large.iterrows():
        start = parse_time(row['应用商店大文件下载开始时间'])
        end = parse_time(row['应用商店大文件下载完成时间'])
        if start is None or end is None:
            continue
        dur = compute_duration(start, end)
        rows.append({
            '运营商': row['运营商'],
            '业务类型': '应用商店大文件下载',
            '开始时间': start,
            '结束时间': end,
            '速率类指标': '应用商店大文件下载速率(Mbps)',
            '数值': round(float(row['应用商店大文件下载速率(Mbps)']), 3) if pd.notna(row['应用商店大文件下载速率(Mbps)']) else 0,
            '时长类指标': '业务时长(秒)',
            '数值.1': dur,
        })
        cnt += 1
    print(f"  提取 {cnt} 行")

    # 构建DataFrame
    result = pd.DataFrame(rows)
    print(f"\n总行数: {len(result)}")

    # 按时间升序排列
    result = result.sort_values('开始时间', ascending=True).reset_index(drop=True)

    # 格式化时间输出
    result['开始时间'] = result['开始时间'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3] if x else '')
    result['结束时间'] = result['结束时间'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3] if x else '')

    # 输出统计
    print(f"\n=== 业务类型分布 ===")
    for biz, cnt in result['业务类型'].value_counts().items():
        ops = result[result['业务类型'] == biz]['运营商'].value_counts()
        op_str = ', '.join([f"{op}: {cnt}" for op, cnt in ops.items()])
        print(f"  {biz}: {cnt} ({op_str})")

    # 时长范围统计
    print(f"\n=== 时长范围统计 ===")
    for biz in result['业务类型'].unique():
        sub = result[result['业务类型'] == biz]
        try:
            vals = sub['数值.1'].astype(float)
            print(f"  {biz}: {min(vals):.3f}s ~ {max(vals):.3f}s (均值{sum(vals)/len(vals):.3f}s)")
        except:
            pass

    # 保存
    cols = ['运营商', '业务类型', '开始时间', '结束时间', '速率类指标', '数值', '时长类指标', '数值.1']
    result = result[cols]
    result.to_excel(OUTPUT_FILE, index=False)
    print(f"\n已保存: {OUTPUT_FILE}")

    # 验证数据完整性
    print(f"\n=== 完整性验证 ===")
    input_ftp_dl = len(data[(data['测试业务']=='FTPDownload') & (data['发起时间'].notna())])
    input_ftp_ul = len(data[(data['测试业务']=='FTPUpload') & (data['发起时间'].notna())])
    input_wx_small = len(data[data['微信小包发送速率(Mbps)'].notna()])
    input_wx_large = len(data[data['微信大包发送速率(Mbps)'].notna()])
    input_app_small = len(data[data['应用商店小文件下载速率(Mbps)'].notna()])
    input_app_large = len(data[data['应用商店大文件下载速率(Mbps)'].notna()])

    output_biz = result['业务类型'].value_counts().to_dict()
    checks = [
        ('FTP下载', input_ftp_dl, output_biz.get('FTP下载', 0), '输入FTPDownload行数（排除空发起时间）'),
        ('FTP上传', input_ftp_ul, output_biz.get('FTP上传', 0), '输入FTPUpload行数（排除空发起时间）'),
        ('微信小包发送', input_wx_small, output_biz.get('微信小包发送', 0), '输入微信小包发送速率非空行数'),
        ('微信大包发送', input_wx_large, output_biz.get('微信大包发送', 0), '输入微信大包发送速率非空行数'),
        ('应用商店小文件下载', input_app_small, output_biz.get('应用商店小文件下载', 0), '输入应用商店小文件速率非空行数'),
        ('应用商店大文件下载', input_app_large, output_biz.get('应用商店大文件下载', 0), '输入应用商店大文件速率非空行数'),
    ]
    all_ok = True
    for biz, inp, out, desc in checks:
        status = '✓' if inp == out else '✗'
        if inp != out:
            all_ok = False
        print(f"  {status} {biz}: 输入{inp}行 → 输出{out}行 ({desc})")
    if all_ok:
        print("\n✓ 所有业务类型行数完全匹配！")
    else:
        print("\n✗ 存在行数不匹配，请检查")

    return result

if __name__ == '__main__':
    result = main()