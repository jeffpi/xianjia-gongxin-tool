"""
Excel导出模块
V2.0.0 全新重写
"""

import os
from typing import Dict, List, Optional, Tuple
import pandas as pd
import numpy as np
from datetime import datetime

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from src.utils.config import APP_NAME, APP_VERSION, OUTPUT_CONFIG


class ExcelExporter:
    """Excel导出器"""

    def __init__(self):
        self.output_config = OUTPUT_CONFIG

    def export(self, results: Dict[str, List[Dict]],
               pending_segments: List[Dict],
               input_files: List[str],
               scene_type: str,
               output_path: str) -> str:
        """
        导出Excel文件

        Args:
            results: 按运营商分组的结果 {运营商: 结果列表}
            pending_segments: 待确认片段列表
            input_files: 输入文件列表
            scene_type: 场景类型
            output_path: 输出路径

        Returns:
            str: 输出文件路径
        """
        # 创建Excel writer
        writer = pd.ExcelWriter(output_path, engine='openpyxl')

        # 导出各运营商sheet
        total_businesses = 0
        for operator, operator_results in results.items():
            if len(operator_results) > 0:
                df = pd.DataFrame(operator_results)
                sheet_name = self.output_config["sheets"].get(operator, operator)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                total_businesses += len(operator_results)

        # 导出汇总sheet（如果有多个运营商）
        if len(results) > 1:
            all_results = []
            for operator_results in results.values():
                all_results.extend(operator_results)
            if all_results:
                df_summary = pd.DataFrame(all_results)
                df_summary.to_excel(writer, sheet_name="汇总", index=False)

        # 导出待确认sheet
        if pending_segments:
            df_pending = self._format_pending_segments(pending_segments)
            df_pending.to_excel(writer, sheet_name="待确认", index=False)

        # 导出版本信息sheet
        version_info = self._create_version_info(
            input_files, scene_type, total_businesses, len(pending_segments)
        )
        df_version = pd.DataFrame(version_info)
        df_version.to_excel(writer, sheet_name="版本信息", index=False)

        # 保存
        writer.close()

        return output_path

    def _format_pending_segments(self, segments: List[Dict]) -> pd.DataFrame:
        """格式化待确认片段"""
        formatted = []

        for seg in segments:
            start_time = seg.get("start_time")
            end_time = seg.get("end_time")

            row = {
                "开始时间": start_time.strftime("%H:%M:%S.%f")[:-3] if hasattr(start_time, 'strftime') else str(start_time),
                "结束时间": end_time.strftime("%H:%M:%S.%f")[:-3] if hasattr(end_time, 'strftime') else str(end_time),
                "时长(ms)": seg.get("duration_ms", 0),
                "候选类型": ", ".join([f"{c[0]}({c[1]:.2f})" for c in seg.get("candidates", [])]),
                "置信度": round(seg.get("confidence", 0), 3),
            }
            formatted.append(row)

        return pd.DataFrame(formatted)

    def _create_version_info(self, input_files: List[str],
                             scene_type: str,
                             business_count: int,
                             pending_count: int) -> List[Dict]:
        """创建版本信息"""
        return [
            {"项目": "工具名称", "值": APP_NAME},
            {"项目": "版本号", "值": f"V{APP_VERSION}"},
            {"项目": "导出时间", "值": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
            {"项目": "场景类型", "值": scene_type},
            {"项目": "输入文件", "值": ", ".join(input_files)},
            {"项目": "正式业务数", "值": business_count},
            {"项目": "待确认片段数", "值": pending_count},
        ]

    def get_default_output_path(self) -> str:
        """获取默认输出路径"""
        timestamp = datetime.now().strftime(self.output_config["timestamp_format"])
        filename = self.output_config["default_filename"].format(timestamp=timestamp)
        return os.path.join(os.getcwd(), filename)

    @staticmethod
    def format_time(dt: datetime) -> str:
        """格式化时间"""
        if dt is None:
            return ""
        return dt.strftime("%H:%M:%S.%f")[:-3]
