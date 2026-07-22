# 5G用户级公共监控速率统计工具 V2.0.0

## 项目结构

```
先甲工信部工具/
├── src/
│   ├── core/          # 核心业务逻辑
│   │   ├── operator_detector.py      # 运营商识别
│   │   ├── scene_inferencer.py       # 场景推断
│   │   ├── business_identifier.py    # 业务识别引擎
│   │   ├── confidence_calculator.py  # 置信度计算
│   │   └── result_generator.py       # 结果生成
│   ├── adapters/      # 数据适配器
│   │   ├── base.py                   # IDataProvider接口
│   │   ├── mmf_reader.py             # 原生MMF解析
│   │   ├── excel_adapter.py          # Excel适配器
│   │   └── xlsb_adapter.py           # XLSB适配器
│   ├── gui/           # 图形界面
│   │   ├── main_window.py            # 主窗口
│   │   ├── import_page.py            # 导入页
│   │   ├── analysis_page.py          # 识别跟踪页
│   │   ├── result_page.py            # 结果导出页
│   │   ├── about_dialog.py           # 关于对话框
│   │   └── version_history.py        # 版本历史
│   └── utils/         # 工具函数
│       ├── config.py                 # 配置与版本号
│       ├── logger.py                 # 日志
│       └── excel_exporter.py         # Excel导出
├── tests/             # 测试
│   ├── test_adapters.py
│   ├── test_operator_detector.py
│   ├── test_scene_inferencer.py
│   ├── test_business_identifier.py
│   └── test_confidence_calculator.py
├── docs/
│   ├── superpowers/
│   │   └── specs/
│   │       └── 2026-07-10-5g-monitor-v2-design.md
│   └── 历史代码文件/   # 旧版本备份
├── 需求文档.md
├── 代码更新日志.md
└── README.md
```

## 模块说明

### adapters（数据适配层）
负责将不同格式的输入文件统一转换为内部标准DataFrame格式。

### core（核心业务层）
- operator_detector：从文件名、数据列推断运营商
- scene_inferencer：扫描整体数据判断场景类型
- business_identifier：综合多维度判断业务类型
- confidence_calculator：多维度评分计算置信度
- result_generator：将识别结果转为统一输出结构

### gui（图形界面层）
三页向导结构：导入→识别跟踪→结果导出

### utils（工具层）
- config：集中管理版本号、业务规格参数
- logger：统一日志记录
- excel_exporter：生成Excel工作簿

## 设计原则

1. 不参考旧代码实现
2. 模块职责清晰，单一职责
3. 接口统一，便于扩展
4. 配置集中管理
5. 版本号单一来源

---

*创建时间：2026-07-10*
*版本：V2.0.0*