"""
业务识别引擎 - V2.0.0修复版
从MMF数据自动识别业务，不依赖基准文件
"""

import os
from typing import Dict, List, Optional, Tuple
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from src.utils.config import BUSINESS_SPECS, INTERVAL_SPECS


class BusinessIdentifier:
    """业务识别引擎 - 从MMF数据自动识别业务"""

    def __init__(self):
        self.business_specs = BUSINESS_SPECS
        self.interval_specs = INTERVAL_SPECS
        
        # 识别阈值
        self.dl_rate_threshold = 10.0  # Mbps，下行高速阈值
        self.ul_rate_threshold = 10.0  # Mbps，上行高速阈值
        self.min_ftp_duration = 5.0    # 秒，FTP业务最小时长
        self.max_ftp_duration = 15.0   # 秒，FTP业务最大时长

    def identify_businesses(self, df: pd.DataFrame, scene_type: str) -> List[Dict]:
        """
        从MMF数据自动识别所有业务片段
        
        Args:
            df: 数据DataFrame（已预处理，包含timestamp、dl_mac、ul_mac等）
            scene_type: 场景类型（暂不使用，自动识别）
            
        Returns:
            list: 业务片段列表
        """
        if len(df) == 0:
            print("  [识别] 数据为空，无法识别")
            return []
        
        # 检查必需列
        required_cols = ["timestamp", "dl_mac", "ul_mac"]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            print(f"  [识别] 缺少必需列: {missing}")
            return []
        
        print(f"  [识别] 开始识别，数据范围: {df['timestamp'].min()} ~ {df['timestamp'].max()}")
        print(f"  [识别] 数据行数: {len(df)}")
        
        # 按时间排序
        df = df.sort_values("timestamp").reset_index(drop=True)
        
        # 识别高速段
        businesses = []
        
        # 1. 识别下行高速段（FTP下载、商店下载）
        dl_segments = self._find_download_segments(df)
        print(f"  [识别] 下行高速段: {len(dl_segments)} 个")
        
        # 2. 识别上行高速段（FTP上传、微信发送）
        ul_segments = self._find_upload_segments(df)
        print(f"  [识别] 上行高速段: {len(ul_segments)} 个")
        
        # 3. 分类和组装业务
        for seg in dl_segments:
            biz = self._classify_download_segment(df, seg)
            if biz:
                businesses.append(biz)
        
        for seg in ul_segments:
            biz = self._classify_upload_segment(df, seg)
            if biz:
                businesses.append(biz)
        
        # 按时间排序
        businesses.sort(key=lambda x: x["start_time"])
        
        print(f"  [识别] 最终识别: {len(businesses)} 个业务")
        
        return businesses

    def _find_download_segments(self, df: pd.DataFrame) -> List[Dict]:
        """查找下行高速段"""
        segments = []
        
        # 标记高速下行
        high_dl = df["dl_mac"] > self.dl_rate_threshold
        
        # 找连续段
        in_segment = False
        start_idx = None
        start_time = None
        
        for i, (is_high, row) in enumerate(zip(high_dl, df.itertuples())):
            if is_high and not in_segment:
                # 开始新段
                in_segment = True
                start_idx = i
                start_time = row.timestamp
            elif not is_high and in_segment:
                # 结束段
                in_segment = False
                end_time = row.timestamp
                duration = (end_time - start_time).total_seconds()
                
                # 只保留持续>1秒的段
                if duration > 1.0:
                    segments.append({
                        "start_idx": start_idx,
                        "end_idx": i - 1,
                        "start_time": start_time,
                        "end_time": end_time,
                        "duration_sec": duration,
                        "direction": "下行",
                    })
        
        # 处理最后一个段
        if in_segment:
            end_time = df.iloc[-1]["timestamp"]
            duration = (end_time - start_time).total_seconds()
            if duration > 1.0:
                segments.append({
                    "start_idx": start_idx,
                    "end_idx": len(df) - 1,
                    "start_time": start_time,
                    "end_time": end_time,
                    "duration_sec": duration,
                    "direction": "下行",
                })
        
        return segments

    def _find_upload_segments(self, df: pd.DataFrame) -> List[Dict]:
        """查找上行高速段"""
        segments = []
        
        # 标记高速上行
        high_ul = df["ul_mac"] > self.ul_rate_threshold
        
        # 找连续段
        in_segment = False
        start_idx = None
        start_time = None
        
        for i, (is_high, row) in enumerate(zip(high_ul, df.itertuples())):
            if is_high and not in_segment:
                # 开始新段
                in_segment = True
                start_idx = i
                start_time = row.timestamp
            elif not is_high and in_segment:
                # 结束段
                in_segment = False
                end_time = row.timestamp
                duration = (end_time - start_time).total_seconds()
                
                # 保留所有上行段（包括短时微信）
                if duration > 0.5:
                    segments.append({
                        "start_idx": start_idx,
                        "end_idx": i - 1,
                        "start_time": start_time,
                        "end_time": end_time,
                        "duration_sec": duration,
                        "direction": "上行",
                    })
        
        # 处理最后一个段
        if in_segment:
            end_time = df.iloc[-1]["timestamp"]
            duration = (end_time - start_time).total_seconds()
            if duration > 0.5:
                segments.append({
                    "start_idx": start_idx,
                    "end_idx": len(df) - 1,
                    "start_time": start_time,
                    "end_time": end_time,
                    "duration_sec": duration,
                    "direction": "上行",
                })
        
        return segments

    def _classify_download_segment(self, df: pd.DataFrame, segment: Dict) -> Optional[Dict]:
        """
        分类下行段为具体业务
        
        规则：
        - 时长5-15秒 + 高速率 -> FTP下载
        - 其他 -> 商店下载（大文件或小文件）
        """
        start_idx = segment["start_idx"]
        end_idx = segment["end_idx"]
        duration = segment["duration_sec"]
        
        # 提取段内数据
        seg_df = df.iloc[start_idx:end_idx+1]
        
        # 计算平均速率（去零）
        rates = seg_df["dl_mac"].dropna()
        rates = rates[rates > 0]
        
        if len(rates) == 0:
            return None
        
        avg_rate = rates.mean()
        
        # 分类规则
        if self.min_ftp_duration <= duration <= self.max_ftp_duration:
            # FTP下载：时长5-15秒
            biz_type = "ftp_download"
            biz_name = "FTP下载"
        elif duration < 5:
            # 商店小文件：短时长高速下载
            biz_type = "store_small"
            biz_name = "商店小文件下载"
        else:
            # 商店大文件：长时长下载
            biz_type = "store_large"
            biz_name = "商店大文件下载"
        
        return {
            "type": biz_type,
            "name": biz_name,
            "start_time": segment["start_time"],
            "end_time": segment["end_time"],
            "duration_sec": duration,
            "duration_ms": int(duration * 1000),
            "direction": "下行",
            "avg_rate": avg_rate,
        }

    def _classify_upload_segment(self, df: pd.DataFrame, segment: Dict) -> Optional[Dict]:
        """
        分类上行段为具体业务
        
        规则：
        - 时长5-15秒 -> FTP上传
        - 时长<5秒 -> 微信发送（小包或大包根据流量判断）
        """
        start_idx = segment["start_idx"]
        end_idx = segment["end_idx"]
        duration = segment["duration_sec"]
        
        # 提取段内数据
        seg_df = df.iloc[start_idx:end_idx+1]
        
        # 计算平均速率（去零）
        rates = seg_df["ul_mac"].dropna()
        rates = rates[rates > 0]
        
        if len(rates) == 0:
            return None
        
        avg_rate = rates.mean()
        
        # 计算流量（如果有TBS列）
        if "ul_tbs" in seg_df.columns:
            tbs = seg_df["ul_tbs"].dropna()
            if len(tbs) > 1:
                flow_bits = tbs.iloc[-1] - tbs.iloc[0]
                flow_mb = flow_bits / 8 / 1024 / 1024
            else:
                flow_mb = avg_rate * duration / 8  # 估算
        else:
            flow_mb = avg_rate * duration / 8  # 估算
        
        # 分类规则
        if self.min_ftp_duration <= duration <= self.max_ftp_duration:
            # FTP上传：时长5-15秒
            biz_type = "ftp_upload"
            biz_name = "FTP上传"
        elif duration < 5:
            # 微信发送：短时长
            if flow_mb < 50:  # <50MB
                biz_type = "wechat_small"
                biz_name = "微信小包发送"
            else:
                biz_type = "wechat_large"
                biz_name = "微信大包发送"
        else:
            # 长时长上行，可能是微信大包
            biz_type = "wechat_large"
            biz_name = "微信大包发送"
        
        return {
            "type": biz_type,
            "name": biz_name,
            "start_time": segment["start_time"],
            "end_time": segment["end_time"],
            "duration_sec": duration,
            "duration_ms": int(duration * 1000),
            "direction": "上行",
            "avg_rate": avg_rate,
            "flow_mb": flow_mb,
        }
