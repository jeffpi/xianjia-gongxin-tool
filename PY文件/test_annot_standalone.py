"""独立测试标注文件生成（不依赖主程序）"""
import sys, os, traceback
sys.stderr = open(os.devnull, 'w')
import importlib.util

spec = importlib.util.spec_from_file_location('t', '5G用户级公共监控速率统计工具-V1.04.py')
t = importlib.util.module_from_spec(spec)
spec.loader.exec_module(t)

out_path = '/Users/sun/ClaudeCode/先甲工信部工具/test_annot.xlsx'

proc = t.UCMProcessor(t.ConfigManager().config)
params = {
    'rate_column': 'RLC',
    'dl_peak_limit': 900,
    'ul_peak_limit': 160,
    'down_min': 50,
    'up_min': 10,
    'match_mode': 'hybrid',
}

print("Parsing...", flush=True)
proc.parse_mmf(['联通/mmf20260703115336-联通.xlsx'], params)
print("Matching...", flush=True)
call_df, stats = proc.match(t.DEFAULT_PLAN, params)
print("Generating annotated...", flush=True)

# 手动调用
sys.stderr = sys.__stderr__
try:
    path = proc.generate_annotated(
        ['联通/mmf20260703115336-联通.xlsx'],
        t.DEFAULT_PLAN,
        params,
        out_path
    )
    print(f'OK: {path}', flush=True)
    print(f'Size: {os.path.getsize(path)} bytes', flush=True)
except Exception as e:
    traceback.print_exc()
    # 检查输出文件是否存在
    if os.path.exists(out_path):
        print(f'Partial output: {os.path.getsize(out_path)} bytes', flush=True)