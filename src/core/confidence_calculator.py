"""
置信度计算模块
V2.0.0 全新重写
"""

import os
from typing import Dict, List, Optional, Tuple
import pandas as pd
import numpy as np

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from src.utils.config import BUSINESS_SPECS, CONFIDENCE_WEIGHTS, CONFIDENCE_THRESHOLD


class ConfidenceCalculator:
    """置信度计算器"""

    def __init__(self):
        self.weights = CONFIDENCE_WEIGHTS
        self.threshold = CONFIDENCE_THRESHOLD
        self.business_specs = BUSINESS_SPECS

    def calculate_confidence(self, df: pd.DataFrame, 
                             segment: Dict,
                             business_type: str) -> float:
        """
        计算业务片段的置信度

        Args:
            df: 数据DataFrame
            segment: 业务片段
            business_type: 业务类型

        Returns:
            float: 置信度 (0-1)
        """
        scores = {}

        # 1. 流量匹配度
        scores["flow_match"] = self._calc_flow_match(df, segment, business_type)

        # 2. 时长接近度
        scores["duration_match"] = self._calc_duration_match(df, segment, business_type)

        # 3. 顺序一致性
        scores["order_consistency"] = self._calc_order_consistency(df, segment, business_type)

        # 4. 组关系合理性
        scores["group_relation"] = self._calc_group_relation(df, segment, business_type)

        # 加权平均
        confidence = sum(
            scores[key] * self.weights[key]
            for key in self.weights
        )

        return confidence

    def _calc_flow_match(self, df: pd.DataFrame, 
                         segment: Dict,
                         business_type: str) -> float:
        """计算流量匹配度"""
        spec = self.business_specs.get(business_type, {})

        # 如果业务不依赖流量规格
        if "flow_min_mb" not in spec or spec["flow_min_mb"] is None:
            return 0.7  # 默认中等置信度

        # 获取段内流量
        start_time = segment.get("start_time")
        end_time = segment.get("end_time")

        if start_time is None or end_time is None:
            return 0.5

        flow_col = "dl_flow" if spec.get("direction") == "下行" else "ul_flow"
        if flow_col not in df.columns:
            return 0.5

        # 提取段内数据
        mask = (df["timestamp"] >= start_time) & (df["timestamp"] <= end_time)
        segment_df = df[mask]

        if len(segment_df) == 0:
            return 0.3

        # 计算流量增量
        flow_start = segment_df[flow_col].iloc[0]
        flow_end = segment_df[flow_col].iloc[-1]
        flow_mb = abs(flow_end - flow_start)

        # 匹配规格
        flow_min = spec.get("flow_min_mb", 0)
        flow_max = spec.get("flow_max_mb", float('inf'))
        typical_flow = spec.get("typical_flow_mb", (flow_min + flow_max) / 2)

        # 计算匹配度
        if flow_min <= flow_mb <= flow_max:
            # 在规格范围内，越接近典型值置信度越高
            if typical_flow > 0:
                deviation = abs(flow_mb - typical_flow) / typical_flow
                return max(0.6, 1.0 - deviation * 0.3)
            return 0.9
        else:
            # 不在规格范围内
            if flow_mb < flow_min:
                deviation = (flow_min - flow_mb) / flow_min
                return max(0.2, 0.6 - deviation)
            else:
                deviation = (flow_mb - flow_max) / flow_max
                return max(0.3, 0.7 - deviation * 0.5)

        return 0.5

    def _calc_duration_match(self, df: pd.DataFrame,
                             segment: Dict,
                             business_type: str) -> float:
        """计算时长接近度"""
        spec = self.business_specs.get(business_type, {})

        # 如果业务不依赖时长规格
        if "typical_duration_sec" not in spec:
            return 0.7

        typical_duration = spec.get("typical_duration_sec", 10)
        actual_duration = segment.get("duration_sec", 0)

        if actual_duration <= 0:
            return 0.3

        # 计算偏差
        deviation = abs(actual_duration - typical_duration) / typical_duration

        # 偏差越小置信度越高
        if deviation <= 0.1:
            return 1.0
        elif deviation <= 0.2:
            return 0.9
        elif deviation <= 0.3:
            return 0.8
        elif deviation <= 0.5:
            return 0.6
        else:
            return 0.4

    def _calc_order_consistency(self, df: pd.DataFrame,
                                segment: Dict,
                                business_type: str) -> float:
        """计算顺序一致性"""
        # 检查业务顺序是否符合预期
        # 例如：商店小文件在大文件前，微信小包在大包前
        
        # 简化实现：默认返回中等置信度
        return 0.7

    def _calc_group_relation(self, df: pd.DataFrame,
                             segment: Dict,
                             business_type: str) -> float:
        """计算组关系合理性"""
        # 检查组间关系
        # 例如：FTP组后有15秒间隔
        
        # 简化实现：默认返回中等置信度
        return 0.7

    def is_confident(self, confidence: float) -> bool:
        """判断置信度是否足够"""
        return confidence >= self.threshold

    def classify_segments(self, df: pd.DataFrame,
                          segments: List[Dict]) -> Tuple[List[Dict], List[Dict]]:
        """
        将片段分类为正式识别和待确认

        Args:
            df: 数据DataFrame
            segments: 业务片段列表

        Returns:
            tuple: (正式片段列表, 待确认片段列表)
        """
        confident_segments = []
        pending_segments = []

        for segment in segments:
            business_type = segment.get("type", "unknown")
            confidence = self.calculate_confidence(df, segment, business_type)
            
            segment["confidence"] = confidence

            if self.is_confident(confidence):
                confident_segments.append(segment)
            else:
                # 生成候选类型
                candidates = self._generate_candidates(df, segment)
                segment["candidates"] = candidates
                pending_segments.append(segment)

        return confident_segments, pending_segments

    def _generate_candidates(self, df: pd.DataFrame, 
                            segment: Dict) -> List[Tuple[str, float]]:
        """生成候选业务类型"""
        candidates = []

        for biz_type, spec in self.business_specs.items():
            confidence = self.calculate_confidence(df, segment, biz_type)
            if confidence > 0.3:  # 最低候选阈值
                candidates.append((biz_type, confidence))

        # 按置信度排序
        candidates.sort(key=lambda x: x[1], reverse=True)

        return candidates[:3]  # 返回前3个候选
