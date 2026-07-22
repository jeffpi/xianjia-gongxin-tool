"""
结果生成模块
V2.0.0 全新重写
"""

import os
from typing import Dict, List, Optional, Tuple
import pandas as pd
import numpy as np
from datetime import datetime

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from src.utils.config import BUSINESS_SPECS


class ResultGenerator:
    """结果生成器"""

    def __init__(self):
        self.business_specs = BUSINESS_SPECS

    def generate_results(self, df: pd.DataFrame,
                         segments: List[Dict],
                         operator: str = "未知") -> List[Dict]:
        """
        生成识别结果

        Args:
            df: 数据DataFrame
            segments: 业务片段列表
            operator: 运营商

        Returns:
            list: 结果列表
        """
        results = []

        # 按时间排序
        sorted_segments = sorted(segments, key=lambda x: x.get("start_time", datetime.min))

        # 分配轮次序号
        round_num = 1
        prev_end_time = None

        for segment in sorted_segments:
            start_time = segment.get("start_time")
            end_time = segment.get("end_time")

            # 计算速率
            avg_rate = self._calculate_average_rate(df, segment)

            # 格式化时间
            start_str = start_time.strftime("%H:%M:%S.%f")[:-3] if hasattr(start_time, 'strftime') else str(start_time)
            end_str = end_time.strftime("%H:%M:%S.%f")[:-3] if hasattr(end_time, 'strftime') else str(end_time)

            result = {
                "轮次序号": round_num,
                "业务类型": segment.get("name", "未知"),
                "开始时间": start_str,
                "结束时间": end_str,
                "时长(ms)": segment.get("duration_ms", 0),
                "平均速率(Mbps)": round(avg_rate, 2),
                "方向": segment.get("direction", "未知"),
                "置信度": round(segment.get("confidence", 0.0), 3),
                "运营商": operator,
            }

            results.append(result)

            # 更新轮次
            if prev_end_time is None:
                prev_end_time = end_time
            else:
                # 如果间隔超过一定时间，认为是新一轮
                if start_time and prev_end_time:
                    gap = (start_time - prev_end_time).total_seconds()
                    if gap > 60:  # 超过60秒认为是新轮次
                        round_num += 1
                    prev_end_time = end_time

        return results

    def _calculate_average_rate(self, df: pd.DataFrame, segment: Dict) -> float:
        """
        计算平均速率

        Args:
            df: 数据DataFrame
            segment: 业务片段

        Returns:
            float: 平均速率（Mbps）
        """
        start_time = segment.get("start_time")
        end_time = segment.get("end_time")
        direction = segment.get("direction", "下行")

        if start_time is None or end_time is None:
            return 0.0

        if "timestamp" not in df.columns:
            return 0.0

        # 提取段内数据
        mask = (df["timestamp"] >= start_time) & (df["timestamp"] <= end_time)
        segment_df = df[mask]

        if len(segment_df) == 0:
            return 0.0

        # 选择速率列
        if direction == "下行":
            rate_col = "dl_rlc" if "dl_rlc" in segment_df.columns else "dl_mac"
        else:
            rate_col = "ul_rlc" if "ul_rlc" in segment_df.columns else "ul_mac"

        if rate_col not in segment_df.columns:
            return 0.0

        # 计算平均速率（去零）
        rates = segment_df[rate_col].dropna()
        rates = rates[rates > 0]

        if len(rates) == 0:
            return 0.0

        return rates.mean()

    def generate_summary(self, results: List[Dict]) -> Dict:
        """
        生成汇总统计

        Args:
            results: 结果列表

        Returns:
            dict: 汇总统计
        """
        summary = {
            "总业务数": len(results),
            "运营商": set(),
            "业务类型分布": {},
            "平均置信度": 0.0,
        }

        if len(results) == 0:
            return summary

        # 统计运营商
        for r in results:
            summary["运营商"].add(r.get("运营商", "未知"))

        # 统计业务类型分布
        for r in results:
            biz_type = r.get("业务类型", "未知")
            summary["业务类型分布"][biz_type] = summary["业务类型分布"].get(biz_type, 0) + 1

        # 计算平均置信度
        confidences = [r.get("置信度", 0) for r in results]
        summary["平均置信度"] = round(sum(confidences) / len(confidences), 3)

        summary["运营商"] = list(summary["运营商"])

        return summary

    def to_dataframe(self, results: List[Dict]) -> pd.DataFrame:
        """
        转换为DataFrame

        Args:
            results: 结果列表

        Returns:
            pd.DataFrame: 结果DataFrame
        """
        if len(results) == 0:
            return pd.DataFrame(columns=[
                "轮次序号", "业务类型", "开始时间", "结束时间",
                "时长(ms)", "平均速率(Mbps)", "方向", "置信度", "运营商"
            ])

        return pd.DataFrame(results)

    def separate_by_operator(self, results: List[Dict]) -> Dict[str, List[Dict]]:
        """
        按运营商分离结果

        Args:
            results: 结果列表

        Returns:
            dict: {运营商: 结果列表}
        """
        separated = {}

        for result in results:
            operator = result.get("运营商", "未知")
            if operator not in separated:
                separated[operator] = []
            separated[operator].append(result)

        return separated
