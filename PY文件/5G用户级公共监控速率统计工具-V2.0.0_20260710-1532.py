#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
5G用户级公共监控速率统计工具 V2.0.0
全新重写版本

启动入口
"""

import os
import sys

# 添加项目路径
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_DIR)

# 导入主窗口
from src.gui.main_window import main

if __name__ == "__main__":
    main()
