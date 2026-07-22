#!/usr/bin/env python3
"""完整测试V4：直接从工具文件加载UCMProcessor"""

import sys, os, warnings, re, types
warnings.filterwarnings('ignore')
import numpy as np
import pandas as pd

sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')

# 读取源文件
with open('/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py', 'r', encoding='utf-8') as f:
    source = f.read()

# 提取需要的类定义（去掉GUI和start_processing之前的部分）
# 找到 UCMProcessor 类
lines = source.split('\n')

# 找到所有类定义位置
class_starts = []
for i, line in enumerate(lines):
    if line.startswith('class '):
        class_starts.append((i, line))

print("Classes found:")
for idx, ln in class_starts:
    print(f"  Line {idx}: {ln}")

# 提取 UCMProcessor 类（含其后的类直到ExcelExporter/ConfigManager）
# 需要from导入已经在文件开头
# 找到UCMProcessor开始
ucm_start = None
for idx, ln in class_starts:
    if ln.startswith('class UCMProcessor'):
        ucm_start = idx
        break

if ucm_start is None:
    print("ERROR: 找不到UCMProcessor类")
    sys.exit(1)

# 找下一个顶层class
next_class = None
for idx, ln in class_starts:
    if idx > ucm_start:
        next_class = idx
        break
if next_class is None:
    next_class = len(lines)

# 也提取前面的import和常量
imports_and_constants = '\n'.join(lines[:ucm_start])

# UCMProcessor 代码
ucm_code = '\n'.join(lines[ucm_start:next_class])

# 去掉对self._filename的引用
ucm_code = ucm_code.replace('self._filename', '"test_file"')

# 组装
full_code = imports_and_constants + '\n\n' + ucm_code + """

# 运行测试
print('\\n=== 测试：联通mmf store_s修复 ===')

config = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10}
proc = UCMProcessor(config)

# 处理联通mmf
agg = proc.parse_mmf('/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx', config)

print(f'秒级聚合后行数: {len(agg)}')

params = {'match_mode': 'auto', 'down_min': 50, 'up_min': 10}
round_results = proc._match_auto(agg, params, 50, 10)

print(f'\\n总轮次: {len(round_results)}')

# 统计各业务
all_rates = {}
for k in ['ftp_dl', 'ftp_ul', 'store_s', 'store_l', 'wx_s', 'wx_l']:
    all_rates[k] = []

for r in round_results:
    for k, s in r.items():
        if k in all_rates:
            all_rates[k].append({
                'rate': s['clip_rate'],
                'dur': s['duration'],
                'flow': s['flow_mb'],
                'start_t': s['start_t'],
                'end_t': s['end_t']
            })

baseline = {
    'ftp_dl': 572.11, 'ftp_ul': 66.39, 'store_s': 176.29,
    'store_l': 1025.84, 'wx_s': 27.81, 'wx_l': 60.48,
}
key_name = {
    'ftp_dl': 'FTP下载', 'ftp_ul': 'FTP上传', 'store_s': '商店小包',
    'store_l': '商店大包', 'wx_s': '微信小文件', 'wx_l': '微信大文件',
}

print('')
for k, values in all_rates.items():
    n = len(values)
    if n == 0:
        print(f'{key_name[k]} N=0')
        continue
    rlc_avg = np.mean([v['rate'] for v in values])
    base = baseline[k]
    diff_pct = (rlc_avg - base) / base * 100
    flag = '✓' if abs(diff_pct) < 15 else ('±' if abs(diff_pct) < 25 else '✗')
    print(f'{key_name[k]} N={n} RLC={rlc_avg:.2f} (基准{base}, {diff_pct:+.1f}%) {flag}')

print(f'\\n总轮次={len(round_results)}')

print('\\n=== store_s 段详情 ===')
for r_idx, r in enumerate(round_results):
    if 'store_s' in r:
        s = r['store_s']
        print(f'轮{r_idx+1}: dur={s["duration"]}s flow={s["flow_mb"]}MB rate={s["clip_rate"]:.1f}')

print('\\n=== 轮次key组成 ===')
for r_idx, r in enumerate(round_results):
    keys = ', '.join(sorted(r.keys()))
    print(f'轮{r_idx+1}: {keys}')

# 验证83秒段被排除
print('\\n=== 验证 ===')
print('所有store_s段时长均<=25s:', all(s['duration'] <= 25 for r in round_results if 'store_s' in r for s in [r['store_s']]))
"""

exec(full_code)