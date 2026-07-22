#!/usr/bin/env python3
"""完整测试：用实际工具UCMProcessor运行，验证store_s修复效果"""

import sys, os, warnings
warnings.filterwarnings('ignore')
import numpy as np

sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')

# 导入实际工具（避免GUI）
from importlib import util
spec = util.spec_from_file_location('tool', '/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py')
mod = util.module_from_spec(spec)

# 只导入UCMProcessor相关类
exec(open('/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py', encoding='utf-8').read().split('class ConfigManager:')[0] +
     'class TestUCM(UCMProcessor):\n    def __init__(self):\n        self.seconds = None\n        self._filename = "test"\n        self.coverage = {}\n        self.params = {}\n        self._last_round_results = []\n    def _fmt_t(self, v):\n        if isinstance(v, pd.Timestamp):\n            return v.strftime("%Y-%m-%d %H:%M:%S")\n        return str(v)\n    def match(self, plan, params):\n        # 仅auto模式\n        agg = self.seconds\n        dl_min = params.get("down_min", 50); ul_min = params.get("up_min", 10)\n        round_results = self._match_auto(agg, params, dl_min, ul_min)\n        self._last_round_results = round_results\n        return round_results\n')

import pandas as pd
# 直接用类的定义
from types import ModuleType

# 直接从工具文件读取UCMProcessor定义
import importlib.util
spec2 = importlib.util.spec_from_file_location('tool2', '/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py')
tool_module = importlib.util.module_from_spec(spec2)
# 加载非GUI部分
with open(spec2.origin, 'r', encoding='utf-8') as f:
    source = f.read()

# 提取UCMProcessor类
# 找到class UCMProcessor: 开始
start = source.find('class UCMProcessor:')
if start < 0:
    print("找不到UCMProcessor")
    sys.exit(1)

# 找到下一个class定义或文件结尾
next_class = source.find('\nclass ', start + 1)
if next_class < 0:
    next_class = len(source)

ucm_source = source[start:next_class]
# 替换依赖
ucm_source = ucm_source.replace('self._filename', '"test"')

# 提取前面的公共代码
prefix = source[:start]
# 去掉GUI相关import
prefix_lines = []
for line in prefix.split('\n'):
    if 'PySide6' in line or 'QApplication' in line or 'QMainWindow' in line:
        continue
    if 'openpyxl' in line and 'Workbook' in line:
        continue
    prefix_lines.append(line)

# 组装可执行的代码
exec_code = '\n'.join(prefix_lines) + '\n' + ucm_source + """

class TestRunner:
    def __init__(self):
        self.processor = None

    def run(self, mmf_path):
        config = {'dl_peak_limit': 900, 'ul_peak_limit': 160, 'down_min': 50, 'up_min': 10}
        self.processor = UCMProcessor(config)
        agg = self.processor.parse_mmf(mmf_path, config)

        params = {'match_mode': 'auto', 'down_min': 50, 'up_min': 10}
        round_results = self.processor._match_auto(agg, params, 50, 10)

        # 统计
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

        return round_results, all_rates

runner = TestRunner()
rounds, rates = runner.run('/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx')

baseline = {
    'ftp_dl': 572.11, 'ftp_ul': 66.39, 'store_s': 176.29,
    'store_l': 1025.84, 'wx_s': 27.81, 'wx_l': 60.48,
}
key_name = {
    'ftp_dl': 'FTP下载', 'ftp_ul': 'FTP上传', 'store_s': '商店小包',
    'store_l': '商店大包', 'wx_s': '微信小文件', 'wx_l': '微信大文件',
}

print('=== 联通(store_s时长修复) ===')
for k, values in rates.items():
    n = len(values)
    if n == 0:
        print(f'{key_name[k]} N=0')
        continue
    rlc_avg = np.mean([v['rate'] for v in values])
    base = baseline[k]
    diff_pct = (rlc_avg - base) / base * 100
    flag = '✓' if abs(diff_pct) < 15 else ('±' if abs(diff_pct) < 25 else '✗')
    print(f'{key_name[k]} N={n} RLC={rlc_avg:.2f} (基准{base}, {diff_pct:+.1f}%) {flag}')

print(f'\n总轮次={len(rounds)}')

print('\n=== store_s 段详情 ===')
for r_idx, r in enumerate(rounds):
    if 'store_s' in r:
        s = r['store_s']
        print(f'轮{r_idx+1}: dur={s["duration"]}s flow={s["flow_mb"]}MB rate={s["clip_rate"]:.1f}')

print('\n=== 轮次详情 ===')
for r_idx, r in enumerate(rounds):
    keys = ', '.join(sorted(r.keys()))
    print(f'轮{r_idx+1}: {keys}')
"""

exec(exec_code)