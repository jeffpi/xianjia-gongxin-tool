"""
场景推断模块
V2.0.0 全新重写

场景类型：
- 标准六业务轮：完整轮次包含FTP下载、FTP上传、商店小文件、商店大文件、微信小包、微信大包
- FTP专项：仅有FTP上传和下载
- 其他专项或不完整测试
"""

import os
from typing import Dict, List, Optional, Tuple
import pandas as pd
import numpy as np

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from src.utils.config import BUSINESS_SPECS


class SceneInferencer:
    """场景推断器"""

    def __init__(self):
        self.business_specs = BUSINESS_SPECS

    def infer_scene(self, df: pd.DataFrame) -> Tuple[str, float, Dict]:
        """
        推断场景类型

        Args:
            df: 数据DataFrame

        Returns:
            tuple: (场景类型, 置信度, 详细信息)
        """
        # 检测各类业务特征
        features = self._detect_business_features(df)

        # 根据特征判断场景
        scene_type, confidence = self._classify_scene(features)

        return scene_type, confidence, features

    def _detect_business_features(self, df: pd.DataFrame) -> Dict:
        """
        检测业务特征

        Args:
            df: 数据DataFrame

        Returns:
            dict: 业务特征字典
        """
        features = {
            "has_ftp_pattern": False,
            "has_store_pattern": False,
            "has_wechat_pattern": False,
            "ftp_segments": 0,
            "store_small_segments": 0,
            "store_large_segments": 0,
            "wechat_small_segments": 0,
            "wechat_large_segments": 0,
            "total_duration_sec": 0,
            "details": {}
        }

        if len(df) == 0:
            return features

        # 计算总时长
        if "timestamp" in df.columns:
            features["total_duration_sec"] = (
                df["timestamp"].max() - df["timestamp"].min()
            ).total_seconds()

        # 检测FTP模式（约10秒的高速段）
        ftp_features = self._detect_ftp_pattern(df)
        features["has_ftp_pattern"] = ftp_features["detected"]
        features["ftp_segments"] = ftp_features["count"]
        features["details"]["ftp"] = ftp_features

        # 检测商店模式（60MB/1100MB下载）
        store_features = self._detect_store_pattern(df)
        features["has_store_pattern"] = store_features["detected"]
        features["store_small_segments"] = store_features["small_count"]
        features["store_large_segments"] = store_features["large_count"]
        features["details"]["store"] = store_features

        # 检测微信模式（5MB/200MB上传）
        wechat_features = self._detect_wechat_pattern(df)
        features["has_wechat_pattern"] = wechat_features["detected"]
        features["wechat_small_segments"] = wechat_features["small_count"]
        features["wechat_large_segments"] = wechat_features["large_count"]
        features["details"]["wechat"] = wechat_features

        return features

    def _detect_ftp_pattern(self, df: pd.DataFrame) -> Dict:
        """检测FTP模式（约10秒高速段）"""
        result = {
            "detected": False,
            "count": 0,
            "download_count": 0,
            "upload_count": 0,
            "segments": []
        }

        # 检查是否有速率数据
        if "dl_rlc" not in df.columns and "dl_mac" not in df.columns:
            return result

        if "ul_rlc" not in df.columns and "ul_mac" not in df.columns:
            return result

        # 简化检测：查找高速段（>10 Mbps）
        dl_col = "dl_rlc" if "dl_rlc" in df.columns else "dl_mac"
        ul_col = "ul_rlc" if "ul_rlc" in df.columns else "ul_mac"

        dl_high = df[dl_col] > 50  # 50 Mbps
        ul_high = df[ul_col] > 10  # 10 Mbps

        # 检测连续高速段
        dl_segments = self._find_continuous_segments(df, dl_high, min_duration_sec=5)
        ul_segments = self._find_continuous_segments(df, ul_high, min_duration_sec=5)

        # FTP段通常在7-15秒
        ftp_dl_segments = [s for s in dl_segments if 7 <= s["duration_sec"] <= 15]
        ftp_ul_segments = [s for s in ul_segments if 7 <= s["duration_sec"] <= 15]

        result["download_count"] = len(ftp_dl_segments)
        result["upload_count"] = len(ftp_ul_segments)
        result["count"] = result["download_count"] + result["upload_count"]
        result["detected"] = result["count"] > 0
        result["segments"] = ftp_dl_segments + ftp_ul_segments

        return result

    def _detect_store_pattern(self, df: pd.DataFrame) -> Dict:
        """检测应用商店模式（60MB/1100MB下载）"""
        result = {
            "detected": False,
            "small_count": 0,
            "large_count": 0,
            "segments": []
        }

        if "dl_flow" not in df.columns:
            return result

        # 检测大流量下载段
        # 通过流量累计判断
        dl_flow = df["dl_flow"].fillna(0)

        # 检测流量增加段
        segments = self._find_flow_segments(df, dl_flow, "dl")

        # 按流量大小分类
        small_segments = [s for s in segments if 40 <= s["flow_mb"] <= 100]
        large_segments = [s for s in segments if 800 <= s["flow_mb"] <= 1500]

        result["small_count"] = len(small_segments)
        result["large_count"] = len(large_segments)
        result["detected"] = result["small_count"] > 0 or result["large_count"] > 0
        result["segments"] = small_segments + large_segments

        return result

    def _detect_wechat_pattern(self, df: pd.DataFrame) -> Dict:
        """检测微信模式（5MB/200MB上传）"""
        result = {
            "detected": False,
            "small_count": 0,
            "large_count": 0,
            "segments": []
        }

        if "ul_flow" not in df.columns:
            return result

        # 检测上传流量段
        ul_flow = df["ul_flow"].fillna(0)
        segments = self._find_flow_segments(df, ul_flow, "ul")

        # 按流量大小分类
        small_segments = [s for s in segments if 2 <= s["flow_mb"] <= 10]
        large_segments = [s for s in segments if 150 <= s["flow_mb"] <= 300]

        result["small_count"] = len(small_segments)
        result["large_count"] = len(large_segments)
        result["detected"] = result["small_count"] > 0 or result["large_count"] > 0
        result["segments"] = small_segments + large_segments

        return result

    def _find_continuous_segments(self, df: pd.DataFrame,
                                   condition: pd.Series,
                                   min_duration_sec: float = 3.0) -> List[Dict]:
        """
        查找连续段

        Args:
            df: 数据DataFrame
            condition: 条件Series（boolean）
            min_duration_sec: 最小持续时长

        Returns:
            list: 段列表
        """
        segments = []

        if "timestamp" not in df.columns:
            return segments

        # 标记变化点
        changes = condition.astype(int).diff().fillna(0)

        # 找到所有段的起止
        starts = df.index[changes == 1].tolist()
        ends = df.index[changes == -1].tolist()

        # 处理边界情况
        if condition.iloc[0]:
            starts.insert(0, df.index[0])
        if condition.iloc[-1]:
            ends.append(df.index[-1])

        # 提取段信息
        for start, end in zip(starts, ends):
            start_time = df.loc[start, "timestamp"]
            end_time = df.loc[end, "timestamp"]
            duration = (end_time - start_time).total_seconds()

            if duration >= min_duration_sec:
                segments.append({
                    "start_idx": start,
                    "end_idx": end,
                    "start_time": start_time,
                    "end_time": end_time,
                    "duration_sec": duration,
                })

        return segments

    def _find_flow_segments(self, df: pd.DataFrame,
                            flow_col: pd.Series,
                            direction: str) -> List[Dict]:
        """
        查找流量段

        Args:
            df: 数据DataFrame
            flow_col: 流量列
            direction: 方向（dl/ul）

        Returns:
            list: 流量段列表
        """
        segments = []

        if "timestamp" not in df.columns:
            return segments

        # 计算流量差分
        flow_diff = flow_col.diff().fillna(0)

        # 找到流量增加的段
        increasing = flow_diff > 0
        segments_info = self._find_continuous_segments(df, increasing, min_duration_sec=1.0)

        # 计算每段流量
        for seg in segments_info:
            start_idx = seg["start_idx"]
            end_idx = seg["end_idx"]

            if isinstance(start_idx, int) and isinstance(end_idx, int):
                flow_increase = flow_col.iloc[end_idx] - flow_col.iloc[start_idx]
                seg["flow_mb"] = abs(flow_increase)
                seg["direction"] = direction
                segments.append(seg)

        return segments

    def _classify_scene(self, features: Dict) -> Tuple[str, float]:
        """
        根据特征分类场景

        Args:
            features: 业务特征

        Returns:
            tuple: (场景类型, 置信度)
        """
        has_ftp = features["has_ftp_pattern"]
        has_store = features["has_store_pattern"]
        has_wechat = features["has_wechat_pattern"]

        # 标准六业务轮：三类业务都有
        if has_ftp and has_store and has_wechat:
            return "标准六业务轮", 0.9

        # FTP专项：只有FTP
        if has_ftp and not has_store and not has_wechat:
            return "FTP专项", 0.85

        # 部分业务场景
        if has_ftp and (has_store or has_wechat):
            return "部分业务轮", 0.7

        # 无法确定
        if has_ftp or has_store or has_wechat:
            return "未知场景", 0.4

        return "未知场景", 0.0