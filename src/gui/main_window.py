"""
主窗口 - 三页向导GUI
V2.0.0 全新重写
"""

import os
import sys
from typing import Dict, List, Optional
import pandas as pd
from datetime import datetime

# GUI框架选择
try:
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QPushButton, QLabel, QStackedWidget, QFileDialog, QMessageBox,
        QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView,
        QComboBox, QLineEdit, QGroupBox, QTextEdit
    )
    from PySide6.QtCore import Qt, QThread, Signal, Slot, QTimer
    from PySide6.QtGui import QFont, QColor
    GUI_FRAMEWORK = "PySide6"
except ImportError:
    from PyQt5.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QPushButton, QLabel, QStackedWidget, QFileDialog, QMessageBox,
        QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView,
        QComboBox, QLineEdit, QGroupBox, QTextEdit
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal as Signal, pyqtSlot as Slot, QTimer
    from PyQt5.QtGui import QFont, QColor
    GUI_FRAMEWORK = "PyQt5"

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from src.utils.config import (
    APP_NAME, APP_VERSION, APP_AUTHOR, APP_CONTACT,
    GUI_CONFIG, COLORS, VERSION_HISTORY
)
from src.utils.logger import AnalysisLogger, app_logger
from src.adapters.base import DataProviderFactory
from src.adapters.excel_adapter import ExcelMMFAdapter
from src.adapters.xlsb_adapter import XLSBAdapter
from src.core.operator_detector import OperatorDetector
from src.core.scene_inferencer import SceneInferencer
from src.core.business_identifier import BusinessIdentifier
from src.core.confidence_calculator import ConfidenceCalculator
from src.core.result_generator import ResultGenerator
from src.utils.excel_exporter import ExcelExporter


class AnalysisThread(QThread):
    """分析线程"""
    
    progress = Signal(str)
    finished = Signal(dict)
    
    def __init__(self, files: List[str]):
        super().__init__()
        self.files = files
        self.logger = AnalysisLogger()
        
    def run(self):
        """执行分析"""
        try:
            results = {
                "success": False,
                "error": None,
                "dataframes": [],
                "operators": [],
                "scene_type": None,
                "businesses": [],
                "pending": [],
            }
            
            # 1. 加载文件
            self.progress.emit("正在加载文件...")
            factory = DataProviderFactory()
            factory.register(ExcelMMFAdapter())
            factory.register(XLSBAdapter())
            
            dataframes = []
            operators = []
            
            for file_path in self.files:
                provider = factory.get_provider(file_path)
                if provider:
                    df = provider.load(file_path)
                    operator, _ = OperatorDetector().detect(file_path, df)
                    
                    dataframes.append(df)
                    operators.append(operator)
                    
                    self.logger.log_file_import(
                        os.path.basename(file_path),
                        provider.get_format_name(),
                        operator
                    )
            
            results["dataframes"] = dataframes
            results["operators"] = operators
            
            # 2. 合并数据
            self.progress.emit("正在合并数据...")
            if len(dataframes) > 1:
                merger = DataMerger()
                merged_df = merger.merge_files(dataframes, operators)
            else:
                merged_df = dataframes[0] if dataframes else pd.DataFrame()
            
            # 3. 场景推断
            self.progress.emit("正在推断场景...")
            scene_inferencer = SceneInferencer()
            scene_type, scene_conf, features = scene_inferencer.infer_scene(merged_df)
            
            self.logger.log_scene_detection(scene_type, scene_conf)
            results["scene_type"] = scene_type
            
            # 4. 业务识别
            self.progress.emit("正在识别业务...")
            identifier = BusinessIdentifier()
            segments = identifier.identify_businesses(merged_df, scene_type)
            
            # 5. 置信度计算
            self.progress.emit("正在计算置信度...")
            calculator = ConfidenceCalculator()
            confident, pending = calculator.classify_segments(merged_df, segments)
            
            # 6. 生成结果
            self.progress.emit("正在生成结果...")
            generator = ResultGenerator()
            
            all_results = []
            for df, operator in zip(dataframes, operators):
                operator_segments = [s for s in confident if s.get("operator", operator) == operator]
                operator_results = generator.generate_results(df, operator_segments, operator)
                all_results.extend(operator_results)
            
            # 按运营商分组
            results_by_operator = {}
            for r in all_results:
                op = r.get("运营商", "未知")
                if op not in results_by_operator:
                    results_by_operator[op] = []
                results_by_operator[op].append(r)
            
            results["businesses"] = results_by_operator
            results["pending"] = pending
            results["success"] = True
            
            self.progress.emit("分析完成")
            self.finished.emit(results)
            
        except Exception as e:
            results["error"] = str(e)
            self.progress.emit(f"错误: {str(e)}")
            self.finished.emit(results)


class ImportPage(QWidget):
    """第一页：导入与预检"""
    
    files_selected = Signal(list)
    
    def __init__(self):
        super().__init__()
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # 标题
        title = QLabel("导入文件")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        layout.addWidget(title)
        
        # 文件选择区
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout(file_group)
        
        self.file_list = QTextEdit()
        self.file_list.setReadOnly(True)
        self.file_list.setMaximumHeight(150)
        file_layout.addWidget(self.file_list)
        
        btn_layout = QHBoxLayout()
        self.select_btn = QPushButton("选择文件")
        self.select_btn.clicked.connect(self.select_files)
        btn_layout.addWidget(self.select_btn)
        
        self.clear_btn = QPushButton("清空")
        self.clear_btn.clicked.connect(self.clear_files)
        btn_layout.addWidget(self.clear_btn)
        
        file_layout.addLayout(btn_layout)
        layout.addWidget(file_group)
        
        # 预检结果区
        preview_group = QGroupBox("预检结果")
        preview_layout = QVBoxLayout(preview_group)
        
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(4)
        self.preview_table.setHorizontalHeaderLabels(["文件名", "格式", "运营商", "行数"])
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        preview_layout.addWidget(self.preview_table)
        
        layout.addWidget(preview_group)
        
        # 下一步按钮
        self.next_btn = QPushButton("下一步：开始识别")
        self.next_btn.setEnabled(False)
        layout.addWidget(self.next_btn)
        
        # 存储文件列表
        self.files = []
        
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择MMF文件", "",
            "MMF文件 (*.xlsx *.xls *.xlsb *.mmf);;所有文件 (*)"
        )
        
        if files:
            self.files = files
            self.file_list.setText("\n".join(files))
            self.update_preview()
            self.next_btn.setEnabled(True)
            
    def clear_files(self):
        self.files = []
        self.file_list.clear()
        self.preview_table.setRowCount(0)
        self.next_btn.setEnabled(False)
        
    def update_preview(self):
        self.preview_table.setRowCount(len(self.files))
        
        for i, file_path in enumerate(self.files):
            self.preview_table.setItem(i, 0, QTableWidgetItem(os.path.basename(file_path)))
            
            ext = os.path.splitext(file_path)[1].lower()
            format_name = {"xlsx": "XLSX", "xls": "XLS", "xlsb": "XLSB", "mmf": "MMF"}.get(ext[1:], ext)
            self.preview_table.setItem(i, 1, QTableWidgetItem(format_name))
            
            # 简化：从文件名推断运营商
            filename = os.path.basename(file_path).lower()
            if "联通" in filename or "unicom" in filename:
                operator = "联通"
            elif "电信" in filename or "telecom" in filename:
                operator = "电信"
            else:
                operator = "未知"
            self.preview_table.setItem(i, 2, QTableWidgetItem(operator))
            
            self.preview_table.setItem(i, 3, QTableWidgetItem("待分析"))
    
    def get_files(self):
        return self.files


class AnalysisPage(QWidget):
    """第二页：识别与跟踪"""
    
    analysis_complete = Signal(dict)
    
    def __init__(self):
        super().__init__()
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # 标题
        title = QLabel("业务识别")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        layout.addWidget(title)
        
        # 控制区
        control_group = QGroupBox("识别控制")
        control_layout = QHBoxLayout(control_group)
        
        self.start_btn = QPushButton("开始识别")
        self.start_btn.clicked.connect(self.start_analysis)
        control_layout.addWidget(self.start_btn)
        
        layout.addWidget(control_group)
        
        # 进度区
        progress_group = QGroupBox("识别进度")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        progress_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("等待开始...")
        progress_layout.addWidget(self.status_label)
        
        layout.addWidget(progress_group)
        
        # 日志区
        log_group = QGroupBox("识别日志")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        layout.addWidget(log_group)
        
        # 下一步按钮
        self.next_btn = QPushButton("下一步：查看结果")
        self.next_btn.setEnabled(False)
        layout.addWidget(self.next_btn)
        
        self.files = []
        self.results = None
        
    def set_files(self, files):
        self.files = files
        
    def start_analysis(self):
        if not self.files:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return
        
        self.start_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        
        # 创建分析线程
        self.analysis_thread = AnalysisThread(self.files)
        self.analysis_thread.progress.connect(self.update_progress)
        self.analysis_thread.finished.connect(self.on_analysis_finished)
        self.analysis_thread.start()
        
    @Slot(str)
    def update_progress(self, message):
        self.log_text.append(message)
        current = self.progress_bar.value()
        self.progress_bar.setValue(min(current + 20, 100))
        self.status_label.setText(message)
        
    @Slot(dict)
    def on_analysis_finished(self, results):
        self.results = results
        self.progress_bar.setValue(100)
        
        if results["success"]:
            self.status_label.setText("分析完成")
            self.next_btn.setEnabled(True)
            
            # 显示汇总
            total = sum(len(v) for v in results["businesses"].values())
            pending = len(results["pending"])
            self.log_text.append(f"\n识别完成: {total} 个业务, {pending} 个待确认片段")
        else:
            self.status_label.setText(f"分析失败: {results['error']}")
        
        self.start_btn.setEnabled(True)
        
    def get_results(self):
        return self.results


class ResultPage(QWidget):
    """第三页：结果与导出"""
    
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.results = None
        self.pending = []
        self.files = []
        self.scene_type = ""
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # 标题
        title = QLabel("识别结果")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        layout.addWidget(title)
        
        # 筛选区
        filter_group = QGroupBox("筛选")
        filter_layout = QHBoxLayout(filter_group)
        
        filter_layout.addWidget(QLabel("运营商:"))
        self.operator_combo = QComboBox()
        self.operator_combo.addItem("全部")
        filter_layout.addWidget(self.operator_combo)
        
        filter_layout.addWidget(QLabel("业务类型:"))
        self.biz_type_combo = QComboBox()
        self.biz_type_combo.addItem("全部")
        filter_layout.addWidget(self.biz_type_combo)
        
        self.operator_combo.currentTextChanged.connect(self.filter_results)
        self.biz_type_combo.currentTextChanged.connect(self.filter_results)
        
        layout.addWidget(filter_group)
        
        # 结果表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(8)
        self.result_table.setHorizontalHeaderLabels([
            "轮次序号", "业务类型", "开始时间", "结束时间",
            "时长(ms)", "平均速率(Mbps)", "方向", "置信度"
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.result_table)
        
        # 导出按钮
        btn_layout = QHBoxLayout()
        
        self.export_btn = QPushButton("导出Excel")
        self.export_btn.clicked.connect(self.export_excel)
        btn_layout.addWidget(self.export_btn)
        
        self.back_btn = QPushButton("返回重新分析")
        btn_layout.addWidget(self.back_btn)
        
        layout.addLayout(btn_layout)
        
    def set_results(self, results, files, scene_type):
        self.results = results.get("businesses", {})
        self.pending = results.get("pending", [])
        self.files = files
        self.scene_type = scene_type
        
        # 更新筛选选项
        operators = list(self.results.keys())
        self.operator_combo.clear()
        self.operator_combo.addItem("全部")
        self.operator_combo.addItems(operators)
        
        biz_types = set()
        for op_results in self.results.values():
            for r in op_results:
                biz_types.add(r.get("业务类型", ""))
        self.biz_type_combo.clear()
        self.biz_type_combo.addItem("全部")
        self.biz_type_combo.addItems(sorted(biz_types))
        
        # 显示结果
        self.display_results()
        
    def display_results(self):
        all_results = []
        for op_results in self.results.values():
            all_results.extend(op_results)
        
        self.result_table.setRowCount(len(all_results))
        
        for i, r in enumerate(all_results):
            self.result_table.setItem(i, 0, QTableWidgetItem(str(r.get("轮次序号", ""))))
            self.result_table.setItem(i, 1, QTableWidgetItem(r.get("业务类型", "")))
            self.result_table.setItem(i, 2, QTableWidgetItem(r.get("开始时间", "")))
            self.result_table.setItem(i, 3, QTableWidgetItem(r.get("结束时间", "")))
            self.result_table.setItem(i, 4, QTableWidgetItem(str(r.get("时长(ms)", ""))))
            self.result_table.setItem(i, 5, QTableWidgetItem(str(r.get("平均速率(Mbps)", ""))))
            self.result_table.setItem(i, 6, QTableWidgetItem(r.get("方向", "")))
            self.result_table.setItem(i, 7, QTableWidgetItem(str(r.get("置信度", ""))))
    
    def filter_results(self):
        operator = self.operator_combo.currentText()
        biz_type = self.biz_type_combo.currentText()
        
        filtered = []
        for op, op_results in self.results.items():
            if operator == "全部" or op == operator:
                for r in op_results:
                    if biz_type == "全部" or r.get("业务类型") == biz_type:
                        filtered.append(r)
        
        self.result_table.setRowCount(len(filtered))
        for i, r in enumerate(filtered):
            self.result_table.setItem(i, 0, QTableWidgetItem(str(r.get("轮次序号", ""))))
            self.result_table.setItem(i, 1, QTableWidgetItem(r.get("业务类型", "")))
            self.result_table.setItem(i, 2, QTableWidgetItem(r.get("开始时间", "")))
            self.result_table.setItem(i, 3, QTableWidgetItem(r.get("结束时间", "")))
            self.result_table.setItem(i, 4, QTableWidgetItem(str(r.get("时长(ms)", ""))))
            self.result_table.setItem(i, 5, QTableWidgetItem(str(r.get("平均速率(Mbps)", ""))))
            self.result_table.setItem(i, 6, QTableWidgetItem(r.get("方向", "")))
            self.result_table.setItem(i, 7, QTableWidgetItem(str(r.get("置信度", ""))))
    
    def export_excel(self):
        output_path, _ = QFileDialog.getSaveFileName(
            self, "保存Excel", 
            ExcelExporter().get_default_output_path(),
            "Excel文件 (*.xlsx)"
        )
        
        if output_path:
            exporter = ExcelExporter()
            exporter.export(
                self.results,
                self.pending,
                [os.path.basename(f) for f in self.files],
                self.scene_type,
                output_path
            )
            
            QMessageBox.information(self, "成功", f"已导出到:\n{output_path}")


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.setup_ui()
        
    def setup_ui(self):
        # 窗口设置
        self.setWindowTitle(GUI_CONFIG["title"])
        self.setMinimumSize(GUI_CONFIG["min_width"], GUI_CONFIG["min_height"])
        self.resize(GUI_CONFIG["window_width"], GUI_CONFIG["window_height"])
        
        # 中央部件
        central = QWidget()
        self.setCentralWidget(central)
        
        layout = QVBoxLayout(central)
        
        # 标题栏
        header = QLabel(f"{APP_NAME} V{APP_VERSION}")
        header.setFont(QFont("Arial", 18, QFont.Bold))
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)
        
        # 页面切换器
        self.pages = QStackedWidget()
        
        # 创建三页
        self.import_page = ImportPage()
        self.analysis_page = AnalysisPage()
        self.result_page = ResultPage()
        
        self.pages.addWidget(self.import_page)
        self.pages.addWidget(self.analysis_page)
        self.pages.addWidget(self.result_page)
        
        layout.addWidget(self.pages)
        
        # 连接信号
        self.import_page.next_btn.clicked.connect(self.go_to_analysis)
        self.analysis_page.next_btn.clicked.connect(self.go_to_result)
        self.result_page.back_btn.clicked.connect(self.go_to_import)
        
    def go_to_analysis(self):
        files = self.import_page.get_files()
        if not files:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return
        
        self.analysis_page.set_files(files)
        self.pages.setCurrentIndex(1)
        
    def go_to_result(self):
        results = self.analysis_page.get_results()
        if not results or not results.get("success"):
            QMessageBox.warning(self, "提示", "请先完成分析")
            return
        
        self.result_page.set_results(
            results,
            self.analysis_page.files,
            results.get("scene_type", "未知")
        )
        self.pages.setCurrentIndex(2)
        
    def go_to_import(self):
        self.pages.setCurrentIndex(0)


def main():
    app = QApplication(sys.argv)
    
    # 设置应用样式
    app.setStyle('Fusion')
    
    # 创建主窗口
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
