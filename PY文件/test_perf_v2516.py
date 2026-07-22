#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
性能测试脚本 V2.5.16 - 最简化版，直接输出到终端
测试核心问题：process() 内部各阶段耗时
"""
import sys, os, time, warnings, importlib.util
from datetime import datetime
import io

# 重定向 stderr 以抑制警告
sys.stderr = io.StringIO()

import pandas as pd
import numpy as np

# 加载主脚本（stderr 已重定向，警告不会显示）
spec = importlib.util.spec_from_file_location(
    "main_tool",
    "/Users/sun/ClaudeCode/先甲工信部工具/PY文件/用户级数据、语音跟踪统计工具_V2.5.16_20260720-0100.py"
)
main_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(main_module)
process = main_module.process

# 文件列表
FILES = [
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/0719-光环/用户级公共-信令跟踪导出_20260718205050202/联通集团测试-ViNR微信视频-被叫-560/mmf20260719114250.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/0719-光环/用户级公共-信令跟踪导出_20260718205050202/联通集团测试-VoNR主叫-563/mmf20260719114339.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/0719-光环/用户级公共-信令跟踪导出_20260718205050202/联通集团测试-数据-553/mmf20260719114359.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/0719-光环/用户级公共-信令跟踪导出_20260718205050202/联通集团测试-ViNR微信视频-主叫-884/mmf20260719114422.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/0719-光环/用户级公共-信令跟踪导出_20260718205050202/光环578/mmf20260719114452.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191205/虚拟用户-信令跟踪导出_20260718205006645/光环578_用户跟踪ID=578/tmf20260719111232/SaveMsg_20260719111232_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191205/虚拟用户-信令跟踪导出_20260718205006645/联通集团测试-数据-553_用户跟踪ID=553/tmf20260719111059/SaveMsg_20260719111059_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191205/虚拟用户-信令跟踪导出_20260718205006645/联通集团测试-数据-553_用户跟踪ID=553/tmf20260719111059/SaveMsg_20260719111144_1.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-ViNR微信视频-被叫-560_用户跟踪ID=560/tmf20260719182020/SaveMsg_20260719182020_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-ViNR微信视频-被叫-560_用户跟踪ID=560/tmf20260719182020/SaveMsg_20260719182136_1.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-VoNR主叫-563_用户跟踪ID=563/tmf20260719181954/SaveMsg_20260719181954_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-VoNR主叫-563_用户跟踪ID=563/tmf20260719181954/SaveMsg_20260719182056_1.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-VoNR被叫-5479_用户跟踪ID=5479/tmf20260719181506/SaveMsg_20260719181506_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-VoNR被叫-5479_用户跟踪ID=5479/tmf20260719181506/SaveMsg_20260719181554_1.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-ViNR微信视频-主叫-884_用户跟踪ID=884/tmf20260719181908/SaveMsg_20260719181908_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-ViNR微信视频-主叫-884_用户跟踪ID=884/tmf20260719181908/SaveMsg_20260719182014_1.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-数据-553_用户跟踪ID=553/tmf20260719181830/SaveMsg_20260719181830_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/虚拟用户-信令跟踪导出_20260719163538384/联通集团测试-数据-553_用户跟踪ID=553/tmf20260719181830/SaveMsg_20260719181921_1.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/用户级-联通集团测试-数据-553_用户级公共监控_TID_15355_20260719163639/mmf20260719182735.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/用户级-联通集团测试-数据-553_用户级公共监控_TID_15355_20260719163639/ALL_FTP_DETAIL_20260719062950.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/07191845/重庆北站/用户级-联通集团测试-数据-553_用户级公共监控_TID_15355_20260719163639/mmf20260719182728.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/用户级-信令跟踪导出_20260719183605658/联通集团测试194-数据新-553/mmf20260719192118.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/用户级-信令跟踪导出_20260719183605658/联通集团测试194-数据-553/mmf20260719192139.xlsx',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/虚拟用户-信令跟踪导出_20260719183525399/联通集团测试194-VoNR被叫-5479_用户跟踪ID=5479/tmf20260719191432/SaveMsg_20260719191432_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/虚拟用户-信令跟踪导出_20260719183525399/联通集团测试194-ViNR微信视频-主叫-884_用户跟踪ID=884/tmf20260719191458/SaveMsg_20260719191458_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/虚拟用户-信令跟踪导出_20260719183525399/联通集团测试194-ViNR微信视频-被叫-560_用户跟踪ID=560/tmf20260719191517/SaveMsg_20260719191517_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/虚拟用户-信令跟踪导出_20260719183525399/联通集团测试194-数据-553_用户跟踪ID=553/tmf20260719191536/SaveMsg_20260719191536_0.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/虚拟用户-信令跟踪导出_20260719183525399/联通集团测试194-数据-553_用户跟踪ID=553/tmf20260719191536/SaveMsg_20260719191626_1.csv',
    '/Users/sun/ClaudeCode/先甲工信部工具/输入文件/联通-0719/最新数据/虚拟用户-信令跟踪导出_20260719183525399/联通集团测试194-VoNR主叫-563_用户跟踪ID=563/tmf20260719191326/SaveMsg_20260719191326_0.csv',
]

valid_files = [f for f in FILES if os.path.exists(f)]
print(f"有效文件: {len(valid_files)} 个", flush=True)

# 强制刷新输出
sys.stdout.flush()

# 回调函数
step_times = []
def time_callback(msg):
    ts = datetime.now().strftime('%H:%M:%S.%f')[:-3]
    step_times.append((ts, msg))
    print(f"[{ts}] {msg}", flush=True)

print("=" * 70, flush=True)
print("开始 process() 完整运行", flush=True)
print("=" * 70, flush=True)

t_start = time.time()
try:
    output_file = process(
        files=valid_files,
        qci_list=[6, 7],
        ftp_duration=20,
        merge_raw=True,
        add_annotations=True,
        callback=time_callback,
        cancel_check=None,
        progress_cb=None,
        base_file=None,
        time_filter=None,
        phone_trace_map=None
    )
    t_end = time.time()
    print(f"\n总耗时: {t_end - t_start:.2f}s", flush=True)
    print(f"输出文件: {output_file}", flush=True)

except Exception as e:
    t_end = time.time()
    print(f"\n错误: {e}", flush=True)
    import traceback
    traceback.print_exc()
    print(f"运行到错误时耗时: {t_end - t_start:.2f}s", flush=True)

# 打印各步骤时间差
if len(step_times) >= 2:
    print("\n" + "=" * 70, flush=True)
    print("各步骤耗时分析", flush=True)
    print("=" * 70, flush=True)
    for i in range(len(step_times) - 1):
        t1 = datetime.strptime(step_times[i][0], '%H:%M:%S.%f')
        t2 = datetime.strptime(step_times[i+1][0], '%H:%M:%S.%f')
        dur = (t2 - t1).total_seconds()
        print(f"{dur:8.2f}s | {step_times[i][1]} -> {step_times[i+1][1]}", flush=True)