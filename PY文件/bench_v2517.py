#!/usr/bin/env python3
"""全量31文件性能基准测试 - 数据业务 + 语音VQI"""
import sys, os, time, traceback
sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具/PY文件')
import importlib.util
spec = importlib.util.spec_from_file_location('tool', '/Users/sun/ClaudeCode/先甲工信部工具/PY文件/用户级数据、语音跟踪统计工具_V2.5.17_20260720-1100.py')
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

# Find all 31 input files
FILES = []
for d in ['输入文件/0719-光环', '输入文件/07191205', '输入文件/07191845', '输入文件/联通-0719']:
    base = os.path.join('/Users/sun/ClaudeCode/先甲工信部工具', d)
    if os.path.exists(base):
        for root, dirs, fns in os.walk(base):
            for fn in fns:
                if fn.endswith(('.xlsx', '.csv')):
                    FILES.append(os.path.join(root, fn))

print(f'找到 {len(FILES)} 个文件\n')
step_times = {}

def timed_log(msg):
    elapsed = time.time() - t_start
    step_times[msg] = elapsed
    print(f'[{elapsed:6.1f}s] {msg}', flush=True)

t_start = time.time()

# ===== 测试1: 数据业务 =====
print('='*60)
print('[数据业务] process() 开始')
print('='*60, flush=True)
try:
    out_data = mod.process(
        files=FILES,
        qci_list=[6, 7],
        ftp_duration=20,
        merge_raw=True,
        add_annotations=True,
        callback=lambda msg: timed_log(msg),
        cancel_check=None,
        progress_cb=None,
        base_file=None,
        time_filter=None,
        phone_trace_map=None
    )
    t_data = time.time() - t_start
    print(f'\n[数据业务] 完成! 耗时 {t_data:.1f}s, 输出: {out_data}', flush=True)
except Exception as e:
    t_data = time.time() - t_start
    print(f'\n[数据业务] 报错! 耗时 {t_data:.1f}s')
    traceback.print_exc()

# ===== 测试2: 语音VQI =====
t2_start = time.time()
print('\n' + '='*60)
print('[语音VQI] process_vqi() 开始')
print('='*60, flush=True)
try:
    out_vqi = mod.process_vqi(
        files=FILES,
        callback=lambda msg: print(f'[{time.time()-t2_start:6.1f}s] [VQI] {msg}', flush=True),
        cancel_check=None,
        progress_cb=None,
        add_annotations=True,
        merge_raw=True,
        phone_trace_map=None
    )
    t_vqi = time.time() - t2_start
    print(f'\n[语音VQI] 完成! 耗时 {t_vqi:.1f}s, 输出: {out_vqi}', flush=True)
except Exception as e:
    t_vqi = time.time() - t2_start
    print(f'\n[语音VQI] 报错! 耗时 {t_vqi:.1f}s')
    traceback.print_exc()

print(f'\n汇总: 数据业务={t_data:.1f}s, 语音VQI={t_vqi:.1f}s', flush=True)
