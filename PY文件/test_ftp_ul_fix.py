#!/usr/bin/env python3
"""测试FTP上传修复结果"""
import sys, os, json, re
sys.path.insert(0, '/Users/sun/ClaudeCode/先甲工信部工具')
import importlib.util
spec = importlib.util.spec_from_file_location('tool', '/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py')
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

mmf = '/Users/sun/ClaudeCode/先甲工信部工具/联通/mmf20260703115336.xlsx'
proc = mod.UCMProcessor(mod.ConfigManager())
params = proc.config.get('global_params', {}).copy()
params['match_mode'] = 'auto'
proc.parse_mmf([mmf], params)
call_df, stats = proc.match(json.loads(json.dumps(proc.config.get('test_plan', []))), params)

sm = stats
print(f"summary type: {type(sm)}")
print(f"summary keys: {list(sm.keys()) if isinstance(sm, dict) else 'N/A'}")

# 模板
templates = {
    'FTP下载': 572.11,
    'FTP上传': 66.39,
    '应用商店_小包': 176.29,
    '应用商店_大包': 1025.84,
    '微信_小文件': 27.81,
    '微信_大文件': 60.48,
}

def get_rate(d, label):
    """获取速率值"""
    for k in ['锟斤拷峰后平均锟斤拷率', '削峰后平均吞吐率', '削峰后上行RLC平均吞吐率(Mbps)', '削峰后下行RLC平均吞吐率(Mbps)', '削峰后下行RLC平均吞吐率']:
        if k in d and d[k] is not None:
            return float(d[k])
    for k, v in d.items():
        print(f"  {k}: {v}")
    return None

# 打印结果
print(f"\n=== 联通(ul_rlc mean+削峰) ===")
total_rounds = sm.get('轮次', None)
for grp, tmpl in templates.items():
    d = sm.get(grp, {})
    # 找到速率值
    rate_val = None
    for k, v in d.items():
        if '削峰' in str(k) and ('平均' in str(k) or '吞吐率' in str(k)) and v is not None:
            rate_val = float(v)
            break
    if rate_val is None:
        for k, v in d.items():
            if '平均吞吐率' in str(k) and v is not None:
                rate_val = float(v)
                break
    if rate_val is None:
        print(f"{grp} 无速率值")
        continue
    diff_pct = (rate_val - tmpl) / tmpl * 100
    abs_diff = abs(diff_pct)
    if abs_diff < 15:
        tag = '✓'
    elif abs_diff < 25:
        tag = '±'
    else:
        tag = '✗'
    print(f"{grp} N={d.get('测试次数', '?')} R={rate_val:.1f}(tmpl={tmpl}, {diff_pct:+.1f}%)[{tag}]")

rounds = sm.get('轮次', sm.get('测试轮次', None))
if rounds is None:
    if call_df is not None:
        rounds = len(call_df) // 2 if len(call_df) > 0 else 0
print(f"总轮次={rounds}")

# 打印详细值
print("\n=== 详细统计 ===")
for grp in ['FTP下载', 'FTP上传', '应用商店_小包', '应用商店_大包', '微信_小文件', '微信_大文件']:
    d = sm.get(grp, {})
    print(f"\n--- {grp} ---")
    for k, v in d.items():
        print(f"  {k}: {v}")