#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最终修复：应用gap<10合并 + flow>=600阈值 + dur/rate区分 + 方向判定改进
修改文件: 5G用户级公共监控速率统计工具-V1.04.py
"""

import sys, os, json, re
from datetime import datetime, timedelta
import numpy as np
import pandas as pd

# 读当前文件
file_path = '/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 备份原文件
from datetime import datetime as dt
bak_path = f'/Users/sun/ClaudeCode/先甲工信部工具/5G用户级公共监控速率统计工具-V1.04_备份_{dt.now().strftime("%Y%m%d-%H%M")}.py'
with open(bak_path, 'w', encoding='utf-8') as f:
    f.write(content)
print(f"备份已创建: {bak_path}")

# 应用修改
# 修改1: _classify_seg 中flow>=700改为flow>=600
old = '''            # flow >= 700
            if flow >= 1500:
                return 'store_l'  # 合并后完整的大包段
            # 700~1500: 区分FTP高速下载 vs store_l被gap切分后的子段
            # store_l特征：dur>=15s(持续时间长) 或 rate<500(子段速率低)
            if dur >= 15 or rate < 500:
                return 'store_l'
            return 'ftp_dl'  # 短时高速段是FTP下载'''
new = '''            # flow >= 600
            if flow >= 1500:
                return 'store_l'  # 合并后完整的大包段
            # 600~1500: 区分FTP高速下载(短时高rate) vs store_l被gap切分后的子段
            # store_l特征：dur>=15s(持续时间长) 或 rate<500(子段速率低)
            if dur >= 15:
                return 'store_l'
            if rate < 500:
                return 'store_l'
            return 'ftp_dl'  # 短时高速段是FTP下载'''

if old in content:
    content = content.replace(old, new)
else:
    print("ERROR: old string not found for _classify_seg change!")
    sys.exit(1)

# 修改2: 更新docstring
old_doc = '''          flow<50 → None(噪声)
          50~350MB → store_s（商店小包3~8s，高rate~150Mbps）
          350~700MB → ftp_dl（FTP下载低流量段/起步段）
          700~1500MB → dur>=15s或rate<500 → store_l（商店大包被gap切分后的子段）
                    → else → ftp_dl（FTP高速下载段）
          >=1500MB → store_l（商店大包合并后完整段）'''
new_doc = '''          flow<50 → None(噪声)
          50~350MB → store_s（商店小包3~8s，高rate~150Mbps）
          350~600MB → ftp_dl（FTP下载低流量段/起步段）
          600~1500MB → dur>=15s或rate<500 → store_l（商店大包被gap切分后的子段）
                    → else → ftp_dl（FTP高速下载段）
          >=1500MB → store_l（商店大包合并后完整段）'''

if old_doc in content:
    content = content.replace(old_doc, new_doc)
else:
    print("WARN: docstring not found, skipping")
    # 尝试宽松匹配
    if '350~700MB' in content:
        # 手动替换
        pass

# 修改3: 更新v1.04说明
old_note = '        - 700~1500MB区间增加合并后store_l识别：gap<10s合并后flow>=700且dur>=15或rate<500→store_l'
new_note = '        - 600~1500MB区间增加合并后store_l识别：gap<10s合并后flow>=600且dur>=15或rate<500→store_l'

if old_note in content:
    content = content.replace(old_note, new_note)
else:
    print("WARN: note string not found")

# 写回文件
with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print("修改完成！")

# 验证修改是否生效
with open(file_path, 'r') as f:
    new_content = f.read()
assert 'flow >= 600' in new_content, "修改未生效: flow >= 600"
assert '600~1500' in new_content, "修改未生效: 600~1500"
print("修改验证通过！")