"""
XLSB MMF 适配器 - 处理.xlsb格式的MMF导出文件
V2.0.0 全新重写
"""

import os
from typing import Dict, List, Optional, Any
import pandas as pd
import numpy as np
from datetime import datetime

from .base import IDataProvider


class XLSBAdapter(IDataProvider):
    """XLSB格式MMF适配器"""

    def __init__(self):
        self._column_mapping = {
            # 时间列
            "Millisecond": "millisecond",
            "采集时间": "timestamp",

            # MAC速率列（bps）
            "Downlink MAC Throughput(bps)": "dl_mac",
            "Uplink MAC Throughput(bps)": "ul_mac",
            "下行MAC吞吐率(bps)": "dl_mac",
            "上行MAC吞吐率(bps)": "ul_mac",

            # RLC速率列（bps）
            "Downlink RLC Throughput(bps)": "dl_rlc",
            "Uplink RLC Throughput(bps)": "ul_rlc",
            "下行RLC吞吐率(bps)": "dl_rlc",
            "上行RLC吞吐率(bps)": "ul_rlc",

            # 流量列
            "Downlink Flow(MB)": "dl_flow",
            "Uplink Flow(MB)": "ul_flow",
            "下行流量(MB)": "dl_flow",
            "上行流量(MB)": "ul_flow",

            # 信号列
            "SSB Beam Rsrp": "rsrp",
            "SRS RSRP": "rsrp",
            "SINR": "sinr",

            # 其他
            "CRNTI": "crnti",
            "CA": "ca",
        }

    def can_handle(self, file_path: str) -> bool:
        """判断是否能处理该文件"""
        ext = os.path.splitext(file_path)[1].lower()
        return ext == ".xlsb"

    def get_format_name(self) -> str:
        """获取格式名称"""
        return "XLSB"

    def load(self, file_path: str) -> pd.DataFrame:
        """
        加载XLSB文件并转换为标准DataFrame

        Args:
            file_path: 文件路径

        Returns:
            pd.DataFrame: 标准格式DataFrame
        """
        # XLSB需要pyxlsb库
        try:
            from pyxlsb import open_workbook
        except ImportError:
            raise ImportError("需要安装pyxlsb库: pip install pyxlsb")

        # 使用pyxlsb读取
        data = []
        with open_workbook(file_path) as wb:
            with wb.get_sheet(0) as sheet:
                for row in sheet.rows():
                    data.append([item.v for item in row])

        # 转换为DataFrame
        df = pd.DataFrame(data)

        # 检测表头
        header_row = self._detect_header_row(df)
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)

        # 标准化列名
        df = self._standardize_columns(df)

        # 单位转换
        df = self._convert_units(df)

        # 处理时间列
        df = self._process_time_column(df)

        return df

    def _detect_header_row(self, df: pd.DataFrame) -> int:
        """检测表头行"""
        keywords = ["MAC", "RLC", "Throughput", "吞吐率", "Millisecond", "采集时间"]

        for idx, row in df.iterrows():
            row_str = " ".join(str(cell) for cell in row if pd.notna(cell))
            if any(kw in row_str for kw in keywords):
                return idx

        return 7

    def _standardize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """标准化列名"""
        new_columns = {}

        for col in df.columns:
            if col in self._column_mapping:
                new_columns[col] = self._column_mapping[col]
            else:
                for key, value in self._column_mapping.items():
                    if key in str(col) or str(col) in key:
                        new_columns[col] = value
                        break

        df = df.rename(columns=new_columns)
        return df

    def _convert_units(self, df: pd.DataFrame) -> pd.DataFrame:
        """单位转换"""
        rate_columns = ["dl_mac", "ul_mac", "dl_rlc", "ul_rlc"]

        for col in rate_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce") / 1_000_000

        return df

    def _process_time_column(self, df: pd.DataFrame) -> pd.DataFrame:
        """处理时间列"""
        if "timestamp" not in df.columns and "millisecond" in df.columns:
            df["timestamp"] = pd.to_datetime(
                pd.to_numeric(df["millisecond"], errors="coerce"),
                unit="ms"
            )

        if "timestamp" in df.columns:
            if not pd.api.types.is_datetime64_any_dtype(df["timestamp"]):
                df["timestamp"] = pd.to_datetime(df["timestamp"])

        return df

    def detect_operator(self, file_path: str, df: pd.DataFrame) -> str:
        """从文件名推断运营商"""
        filename = os.path.basename(file_path).lower()

        if any(kw in filename for kw in ["联通", "unicom", "lt"]):
            return "联通"
        if any(kw in filename for kw in ["电信", "telecom", "ct", "dx"]):
            return "电信"

        return "未知"