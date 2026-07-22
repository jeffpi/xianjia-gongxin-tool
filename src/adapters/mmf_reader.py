"""
原生MMF文件读取器
V2.0.0 全新重写

MMF文件是华为UCM监控数据的原生格式，需要特殊解析。
"""

import os
import struct
from typing import Dict, List, Optional, Any, Tuple
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

from .base import IDataProvider


class MMFReader(IDataProvider):
    """原生MMF文件读取器"""

    def __init__(self):
        # MMF文件头标识
        self.MMFT_MAGIC = b'MMFT'
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

        if ext == ".mmf":
            return True

        # 有些.mmf实际上是Excel格式
        if ext in [".xlsx", ".xls", ".xlsb"]:
            # 检查文件名是否包含mmf
            if "mmf" in file_path.lower():
                return False  # 交给Excel适配器处理

        return False

    def get_format_name(self) -> str:
        """获取格式名称"""
        return "MMF"

    def load(self, file_path: str) -> pd.DataFrame:
        """
        加载MMF文件并转换为标准DataFrame

        Args:
            file_path: 文件路径

        Returns:
            pd.DataFrame: 标准格式DataFrame

        Note:
            原生MMF文件解析复杂，需要：
            1. 解析文件头
            2. 解析数据块
            3. 解析字段定义
            4. 提取数据行

            如果解析失败，建议先导出为Excel格式
        """
        # 实际的MMF文件解析需要华为的解析库或详细的格式文档
        # 这里提供一个框架，实际解析可能需要使用UCM工具导出

        try:
            # 尝试解析MMF文件头
            with open(file_path, 'rb') as f:
                magic = f.read(4)

                if magic == self.MMFT_MAGIC:
                    # 真正的MMF格式，需要完整解析
                    return self._parse_native_mmf(f)
                else:
                    # 可能是已转换的格式，尝试作为文本解析
                    return self._parse_as_text(file_path)

        except Exception as e:
            raise ValueError(f"无法解析MMF文件 {file_path}: {str(e)}")

    def _parse_native_mmf(self, file_obj) -> pd.DataFrame:
        """
        解析原生MMF格式

        这是一个框架实现，实际MMF格式需要详细文档
        """
        # 读取文件版本
        version = struct.unpack('<I', file_obj.read(4))[0]

        # 读取字段定义
        field_count = struct.unpack('<I', file_obj.read(4))[0]
        fields = []
        for _ in range(field_count):
            field_name_len = struct.unpack('<I', file_obj.read(4))[0]
            field_name = file_obj.read(field_name_len).decode('utf-8')
            fields.append(field_name)

        # 读取数据行
        data = []
        while True:
            # 读取行长度
            length_bytes = file_obj.read(4)
            if not length_bytes:
                break
            row_length = struct.unpack('<I', length_bytes)[0]

            # 读取行数据
            row_data = file_obj.read(row_length)
            if not row_data:
                break

            # 解析行数据（根据字段定义）
            row = self._parse_row_data(row_data, fields)
            data.append(row)

        # 转换为DataFrame
        df = pd.DataFrame(data, columns=fields[:len(data[0])] if data else fields)

        # 标准化列名
        df = self._standardize_columns(df)

        # 单位转换
        df = self._convert_units(df)

        return df

    def _parse_row_data(self, data: bytes, fields: List[str]) -> List[Any]:
        """解析单行数据"""
        # 实际解析需要根据MMF格式文档
        # 这里返回空列表作为占位
        return []

    def _parse_as_text(self, file_path: str) -> pd.DataFrame:
        """作为文本文件解析"""
        # 尝试读取为CSV或固定宽度格式
        try:
            # 尝试CSV格式
            df = pd.read_csv(file_path, encoding='utf-8', low_memory=False)
        except:
            try:
                # 尝试GBK编码
                df = pd.read_csv(file_path, encoding='gbk', low_memory=False)
            except:
                # 尝试固定宽度格式
                df = pd.read_fwf(file_path, encoding='utf-8')

        # 标准化列名
        df = self._standardize_columns(df)

        # 单位转换
        df = self._convert_units(df)

        # 处理时间列
        df = self._process_time_column(df)

        return df

    def _standardize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """标准化列名"""
        new_columns = {}

        for col in df.columns:
            col_str = str(col)
            if col_str in self._column_mapping:
                new_columns[col] = self._column_mapping[col_str]
            else:
                for key, value in self._column_mapping.items():
                    if key in col_str or col_str in key:
                        new_columns[col] = value
                        break

        df = df.rename(columns=new_columns)
        return df

    def _convert_units(self, df: pd.DataFrame) -> pd.DataFrame:
        """单位转换：bps -> Mbps"""
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


class MMFConverter:
    """MMF文件转换器（用于导出为Excel）"""

    @staticmethod
    def convert_to_excel(mmf_path: str, output_path: str):
        """
        将MMF文件转换为Excel

        Args:
            mmf_path: MMF文件路径
            output_path: 输出Excel路径
        """
        reader = MMFReader()
        df = reader.load(mmf_path)
        df.to_excel(output_path, index=False)