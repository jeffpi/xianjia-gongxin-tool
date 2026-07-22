"""
日志模块 - 统一日志记录
V2.0.0 全新重写
"""

import logging
import sys
from datetime import datetime
from pathlib import Path

def setup_logger(name: str, log_file: str = None, level=logging.INFO):
    """
    配置并返回logger

    Args:
        name: logger名称
        log_file: 日志文件路径（可选）
        level: 日志级别

    Returns:
        logging.Logger: 配置好的logger
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # 避免重复添加handler
    if logger.handlers:
        return logger

    # 控制台handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)

    # 格式化
    formatter = logging.Formatter(
        '[%(asctime)s] %(levelname)s [%(name)s] %(message)s',
        datefmt='%H:%M:%S'
    )
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 文件handler（如果指定）
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger


def get_logger(name: str = "app"):
    """获取logger实例"""
    return logging.getLogger(name)


# 应用级logger
app_logger = setup_logger("app")


class AnalysisLogger:
    """分析过程日志记录器"""

    def __init__(self):
        self.logs = []
        self.current_round = 0
        self.current_segment = None

    def log_file_import(self, filename: str, format_type: str, operator: str):
        """记录文件导入"""
        self.logs.append({
            "time": datetime.now().isoformat(),
            "type": "file_import",
            "filename": filename,
            "format": format_type,
            "operator": operator,
        })
        app_logger.info(f"导入文件: {filename} (格式:{format_type}, 运营商:{operator})")

    def log_scene_detection(self, scene_type: str, confidence: float):
        """记录场景检测"""
        self.logs.append({
            "time": datetime.now().isoformat(),
            "type": "scene_detection",
            "scene": scene_type,
            "confidence": confidence,
        })
        app_logger.info(f"场景检测: {scene_type} (置信度:{confidence:.2f})")

    def log_round_start(self, round_num: int):
        """记录轮次开始"""
        self.current_round = round_num
        self.logs.append({
            "time": datetime.now().isoformat(),
            "type": "round_start",
            "round": round_num,
        })
        app_logger.info(f"开始分析第 {round_num} 轮")

    def log_segment_identified(self, business_type: str, start_time: str,
                                end_time: str, confidence: float, evidence: dict):
        """记录业务片段识别"""
        self.logs.append({
            "time": datetime.now().isoformat(),
            "type": "segment_identified",
            "round": self.current_round,
            "business": business_type,
            "start": start_time,
            "end": end_time,
            "confidence": confidence,
            "evidence": evidence,
        })
        app_logger.info(
            f"识别业务: {business_type} [{start_time} ~ {end_time}] "
            f"置信度:{confidence:.2f}"
        )

    def log_segment_pending(self, start_time: str, end_time: str,
                            candidates: list):
        """记录待确认片段"""
        self.logs.append({
            "time": datetime.now().isoformat(),
            "type": "segment_pending",
            "round": self.current_round,
            "start": start_time,
            "end": end_time,
            "candidates": candidates,
        })
        app_logger.warning(
            f"待确认片段: [{start_time} ~ {end_time}] "
            f"候选:{', '.join(candidates)}"
        )

    def log_export(self, filename: str, sheet_count: int, business_count: int):
        """记录导出"""
        self.logs.append({
            "time": datetime.now().isoformat(),
            "type": "export",
            "filename": filename,
            "sheets": sheet_count,
            "businesses": business_count,
        })
        app_logger.info(f"导出Excel: {filename} ({sheet_count} sheets, {business_count} 业务)")

    def get_logs(self):
        """获取所有日志"""
        return self.logs

    def clear(self):
        """清空日志"""
        self.logs = []
        self.current_round = 0
        self.current_segment = None