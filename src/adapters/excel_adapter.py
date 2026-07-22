"""
Excel MMF 适配器 - 简化版
"""

import os
import re
import pandas as pd
from datetime import datetime, timedelta

from .base import IDataProvider


class ExcelMMFAdapter(IDataProvider):
    """Excel格式MMF适配器"""

    def __init__(self):
        self._column_mapping = {
            "Time": "time_raw",
            "Downlink MAC Throughput(bps)": "dl_mac",
            "Uplink MAC Throughput(bps)": "ul_mac",
            "Downlink RLC Throughput(bps)": "dl_rlc",
            "Uplink RLC Throughput(bps)": "ul_rlc",
            "Code0 DL TBS Sum(bit)": "dl_tbs",
            "UL TBS Sum(bit)": "ul_tbs",
            "CRNTI": "crnti",
        }

    def can_handle(self, file_path: str) -> bool:
        ext = os.path.splitext(file_path)[1].lower()
        return ext in [".xlsx", ".xls"]

    def get_format_name(self) -> str:
        return "XLSX"

    def load(self, file_path: str) -> pd.DataFrame:
        """加载Excel文件"""
        # 读取Excel
        df = pd.read_excel(file_path, header=0)

        # 标准化列名
        df = self._standardize_columns(df)

        # 解析时间
        df = self._parse_time(df)

        # 单位转换
        df = self._convert_units(df)

        return df

    def _standardize_columns(self, df):
        """标准化列名"""
        new_cols = {}
        for col in df.columns:
            col_str = str(col).strip()
            if col_str in self._column_mapping:
                new_cols[col] = self._column_mapping[col_str]
            else:
                for k, v in self._column_mapping.items():
                    if k in col_str:
                        new_cols[col] = v
                        break
        if new_cols:
            df = df.rename(columns=new_cols)
        return df

    def _parse_time(self, df):
        """解析时间"""
        if 'time_raw' not in df.columns:
            if 'Time' in df.columns:
                df['time_raw'] = df['Time']
            else:
                return df

        def parse(t):
            try:
                t = str(t)
                m = re.match(r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\((\d+)\)', t)
                if m:
                    dt = datetime.strptime(m.group(1), '%Y-%m-%d %H:%M:%S')
                    return dt + timedelta(milliseconds=int(m.group(2)))
                return pd.NaT
            except:
                return pd.NaT

        # 使用列表推导式，避免pandas索引问题
        timestamps = [parse(t) for t in df['time_raw']]

        # 创建新DataFrame
        result = pd.DataFrame({
            'timestamp': timestamps
        })

        # 合并原始数据
        for col in df.columns:
            if col != 'time_raw':
                result[col] = df[col].values

        result = result[pd.notna(result['timestamp'])]
        result = result.reset_index(drop=True)

        return result

    def _convert_units(self, df):
        """单位转换"""
        for col in ['dl_mac', 'ul_mac', 'dl_rlc', 'ul_rlc']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce') / 1e6
        return df

    def detect_operator(self, file_path, df):
        """识别运营商"""
        filename = os.path.basename(file_path).lower()
        if any(kw in filename for kw in ['联通', 'unicom', 'lt']):
            return '联通'
        if any(kw in filename for kw in ['电信', 'telecom', 'ct', 'dx']):
            return '电信'
        return '未知'