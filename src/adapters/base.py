"""
数据适配器基类 - 定义统一接口
V2.0.0 全新重写
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any
import pandas as pd


class IDataProvider(ABC):
    """
    数据提供者接口
    所有数据适配器必须实现此接口
    """

    @abstractmethod
    def can_handle(self, file_path: str) -> bool:
        """
        判断是否能处理该文件

        Args:
            file_path: 文件路径

        Returns:
            bool: 是否能处理
        """
        pass

    @abstractmethod
    def load(self, file_path: str) -> pd.DataFrame:
        """
        加载文件并转换为标准DataFrame

        Args:
            file_path: 文件路径

        Returns:
            pd.DataFrame: 标准格式的DataFrame

        标准列:
            - timestamp: 时间戳（datetime）
            - millisecond: 毫秒数（int）
            - dl_mac: 下行MAC速率（Mbps）
            - ul_mac: 上行MAC速率（Mbps）
            - dl_rlc: 下行RLC速率（Mbps）
            - ul_rlc: 上行RLC速率（Mbps）
            - dl_flow: 下行累计流量（MB）
            - ul_flow: 上行累计流量（MB）
            - crnti: CRNTI标识
            - ca: CA聚合标识
            - rsrp: RSRP（dBm）
            - sinr: SINR（dB）
        """
        pass

    @abstractmethod
    def get_format_name(self) -> str:
        """
        获取格式名称

        Returns:
            str: 格式名称（如 "MMF", "XLSX", "XLSB"）
        """
        pass

    def validate_data(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        验证数据完整性

        Args:
            df: 数据DataFrame

        Returns:
            dict: 验证结果，包含:
                - valid: bool 是否有效
                - missing_columns: list 缺失列
                - row_count: int 行数
                - time_range: tuple 时间范围
                - warnings: list 警告信息
        """
        result = {
            "valid": True,
            "missing_columns": [],
            "row_count": len(df),
            "time_range": None,
            "warnings": [],
        }

        # 检查必需列
        required_columns = ["timestamp", "millisecond"]
        for col in required_columns:
            if col not in df.columns:
                result["missing_columns"].append(col)
                result["valid"] = False

        # 检查时间范围
        if "timestamp" in df.columns and len(df) > 0:
            result["time_range"] = (
                df["timestamp"].min(),
                df["timestamp"].max()
            )

        # 检查数据量
        if len(df) == 0:
            result["warnings"].append("数据为空")

        return result

    def detect_operator(self, file_path: str, df: pd.DataFrame) -> str:
        """
        从文件名或数据内容推断运营商

        Args:
            file_path: 文件路径
            df: 数据DataFrame

        Returns:
            str: 运营商名称（联通/电信/未知）
        """
        import os

        filename = os.path.basename(file_path).lower()

        # 从文件名推断
        if any(kw in filename for kw in ["联通", "unicom", "lt"]):
            return "联通"
        if any(kw in filename for kw in ["电信", "telecom", "ct", "dx"]):
            return "电信"

        return "未知"


class DataProviderFactory:
    """数据提供者工厂"""

    def __init__(self):
        self._providers: List[IDataProvider] = []

    def register(self, provider: IDataProvider):
        """注册数据提供者"""
        self._providers.append(provider)

    def get_provider(self, file_path: str) -> Optional[IDataProvider]:
        """获取能处理该文件的提供者"""
        for provider in self._providers:
            if provider.can_handle(file_path):
                return provider
        return None

    def get_all_formats(self) -> List[str]:
        """获取所有支持的格式"""
        return [p.get_format_name() for p in self._providers]


class DataMerger:
    """多文件数据合并器"""

    def merge_files(self, dataframes: List[pd.DataFrame],
                    operators: List[str]) -> pd.DataFrame:
        """
        合并多个DataFrame

        Args:
            dataframes: DataFrame列表
            operators: 对应的运营商列表

        Returns:
            pd.DataFrame: 合并后的DataFrame
        """
        if not dataframes:
            return pd.DataFrame()

        # 添加运营商列
        for df, operator in zip(dataframes, operators):
            df["_operator"] = operator

        # 按时间排序合并
        merged = pd.concat(dataframes, ignore_index=True)
        merged = merged.sort_values("timestamp").reset_index(drop=True)

        return merged

    def check_time_overlap(self, df: pd.DataFrame) -> List[Dict]:
        """
        检查时间重叠

        Args:
            df: 合并后的DataFrame

        Returns:
            list: 重叠区间列表
        """
        overlaps = []

        if "_operator" not in df.columns:
            return overlaps

        operators = df["_operator"].unique()

        for i, op1 in enumerate(operators):
            for op2 in operators[i+1:]:
                df1 = df[df["_operator"] == op1]
                df2 = df[df["_operator"] == op2]

                if len(df1) == 0 or len(df2) == 0:
                    continue

                start1, end1 = df1["timestamp"].min(), df1["timestamp"].max()
                start2, end2 = df2["timestamp"].min(), df2["timestamp"].max()

                # 检查重叠
                overlap_start = max(start1, start2)
                overlap_end = min(end1, end2)

                if overlap_start < overlap_end:
                    overlaps.append({
                        "operator1": op1,
                        "operator2": op2,
                        "start": overlap_start,
                        "end": overlap_end,
                    })

        return overlaps