"""
核心模块单元测试
V2.0.0
"""

import os
import sys
import unittest
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.utils.config import APP_VERSION, BUSINESS_SPECS, CONFIDENCE_THRESHOLD
from src.core.operator_detector import OperatorDetector
from src.core.scene_inferencer import SceneInferencer
from src.core.business_identifier import BusinessIdentifier
from src.core.confidence_calculator import ConfidenceCalculator
from src.core.result_generator import ResultGenerator


class TestOperatorDetector(unittest.TestCase):
    """运营商识别测试"""
    
    def setUp(self):
        self.detector = OperatorDetector()
    
    def test_detect_unicom_from_filename(self):
        """测试从文件名识别联通"""
        result = self.detector.detect_from_filename("mmf20260703115336-联通.xlsx")
        self.assertEqual(result, "联通")
    
    def test_detect_telecom_from_filename(self):
        """测试从文件名识别电信"""
        result = self.detector.detect_from_filename("mmf20260703115328-电信.xlsx")
        self.assertEqual(result, "电信")
    
    def test_detect_unknown_from_filename(self):
        """测试未知运营商"""
        result = self.detector.detect_from_filename("data.xlsx")
        self.assertEqual(result, "未知")


class TestSceneInferencer(unittest.TestCase):
    """场景推断测试"""
    
    def setUp(self):
        self.inferencer = SceneInferencer()
    
    def test_infer_empty_dataframe(self):
        """测试空数据"""
        df = pd.DataFrame()
        scene_type, confidence, features = self.inferencer.infer_scene(df)
        self.assertEqual(scene_type, "未知场景")


class TestBusinessIdentifier(unittest.TestCase):
    """业务识别测试"""
    
    def setUp(self):
        self.identifier = BusinessIdentifier()
    
    def test_identify_empty_dataframe(self):
        """测试空数据"""
        df = pd.DataFrame()
        results = self.identifier.identify_businesses(df, "FTP专项")
        self.assertEqual(len(results), 0)


class TestConfidenceCalculator(unittest.TestCase):
    """置信度计算测试"""
    
    def setUp(self):
        self.calculator = ConfidenceCalculator()
    
    def test_confidence_threshold(self):
        """测试置信度阈值"""
        self.assertEqual(self.calculator.threshold, 0.6)
    
    def test_is_confident(self):
        """测试置信度判断"""
        self.assertTrue(self.calculator.is_confident(0.7))
        self.assertFalse(self.calculator.is_confident(0.5))


class TestResultGenerator(unittest.TestCase):
    """结果生成测试"""
    
    def setUp(self):
        self.generator = ResultGenerator()
    
    def test_generate_empty_results(self):
        """测试空结果"""
        df = pd.DataFrame()
        segments = []
        results = self.generator.generate_results(df, segments, "联通")
        self.assertEqual(len(results), 0)
    
    def test_generate_summary_empty(self):
        """测试空汇总"""
        results = []
        summary = self.generator.generate_summary(results)
        self.assertEqual(summary["总业务数"], 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
