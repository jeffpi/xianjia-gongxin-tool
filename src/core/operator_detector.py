"""
运营商识别模块
V2.0.0 全新重写
"""

import os
import re
from typing import Dict, List, Optional, Tuple
import pandas as pd
import numpy as np

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from src.utils.config import OPERATOR_KEYWORDS


class OperatorDetector:
    """运营商识别器"""

    def __init__(self):
        self.keywords = OPERATOR_KEYWORDS

    def detect_from_filename(self, filename: str) -> str:
        """
        从文件名推断运营商

        Args:
            filename: 文件名

        Returns:
            str: 运营商名称（联通/电信/未知）
        """
        filename_lower = filename.lower()

        # 检查联通关键词
        for kw in self.keywords["联通"]:
            if kw in filename_lower:
                return "联通"

        # 检查电信关键词
        for kw in self.keywords["电信"]:
            if kw in filename_lower:
                return "电信"

        return "未知"

    def detect_from_data(self, df: pd.DataFrame) -> str:
        """
        从数据内容推断运营商

        Args:
            df: 数据DataFrame

        Returns:
            str: 运营商名称
        """
        # 检查是否有运营商标识列
        for col in df.columns:
            col_lower = str(col).lower()
            if "operator" in col_lower or "运营商" in col_lower:
                values = df[col].dropna().unique()
                if len(values) > 0:
                    return str(values[0])

        # 检查PLMN列
        for col in df.columns:
            col_lower = str(col).lower()
            if "plmn" in col_lower:
                values = df[col].dropna().unique()
                if len(values) > 0:
                    plmn = str(values[0])
                    if plmn.startswith("46001"):
                        return "联通"
                    elif plmn.startswith("46003") or plmn.startswith("46011"):
                        return "电信"

        return "未知"

    def detect(self, file_path: str, df: pd.DataFrame) -> Tuple[str, float]:
        """
        综合推断运营商

        Args:
            file_path: 文件路径
            df: 数据DataFrame

        Returns:
            tuple: (运营商名称, 置信度)
        """
        # 从文件名推断
        filename_op = self.detect_from_filename(file_path)

        # 从数据推断
        data_op = self.detect_from_data(df)

        # 综合判断
        if filename_op != "未知" and data_op != "未知":
            if filename_op == data_op:
                return filename_op, 1.0
            else:
                # 文件名优先级更高
                return filename_op, 0.8

        if filename_op != "未知":
            return filename_op, 0.7

        if data_op != "未知":
            return data_op, 0.7

        return "未知", 0.0

    def batch_detect(self, file_paths: List[str],
                     dataframes: List[pd.DataFrame]) -> Dict[str, str]:
        """
        批量识别运营商

        Args:
            file_paths: 文件路径列表
            dataframes: DataFrame列表

        Returns:
            dict: {文件路径: 运营商}
        """
        results = {}

        for file_path, df in zip(file_paths, dataframes):
            operator, _ = self.detect(file_path, df)
            results[file_path] = operator

        return results

    def group_by_operator(self, file_paths: List[str],
                          dataframes: List[pd.DataFrame]) -> Dict[str, List[Tuple[str, pd.DataFrame]]]:
        """
        按运营商分组

        Args:
            file_paths: 文件路径列表
            dataframes: DataFrame列表

        Returns:
            dict: {运营商: [(文件路径, DataFrame), ...]}
        """
        groups = {
            "联通": [],
            "电信": [],
            "未知": []
        }

        for file_path, df in zip(file_paths, dataframes):
            operator, _ = self.detect(file_path, df)
            groups[operator].append((file_path, df))

        return groups