"""
配置模块 - 集中管理版本号、业务规格参数
V2.0.0 全新重写，不参考旧代码
"""

# ==================== 版本信息 ====================
APP_NAME = "5G用户级公共监控速率统计工具"
APP_VERSION = "2.0.0"
APP_AUTHOR = "孙晓军"
APP_CONTACT = "317827@qq.com"

# ==================== 业务规格 ====================
BUSINESS_SPECS = {
    "ftp_download": {
        "name": "FTP下载",
        "direction": "下行",
        "typical_duration_sec": 10,
        "flow_min_mb": None,  # 不固定流量
        "flow_max_mb": None,
    },
    "ftp_upload": {
        "name": "FTP上传",
        "direction": "上行",
        "typical_duration_sec": 10,
        "flow_min_mb": None,
        "flow_max_mb": None,
    },
    "store_small": {
        "name": "商店小文件下载",
        "direction": "下行",
        "typical_flow_mb": 60,
        "flow_min_mb": 40,
        "flow_max_mb": 100,
        "max_wait_sec": 300,
    },
    "store_large": {
        "name": "商店大文件下载",
        "direction": "下行",
        "typical_flow_mb": 1100,
        "flow_min_mb": 800,
        "flow_max_mb": 1500,
        "max_wait_sec": 300,
    },
    "wechat_small": {
        "name": "微信小包发送",
        "direction": "上行",
        "typical_flow_mb": 5,
        "flow_min_mb": 2,
        "flow_max_mb": 10,
    },
    "wechat_large": {
        "name": "微信大包发送",
        "direction": "上行",
        "typical_flow_mb": 200,
        "flow_min_mb": 150,
        "flow_max_mb": 300,
    },
}

# ==================== 时间间隔规格 ====================
INTERVAL_SPECS = {
    "ftp_internal_sec": 5,        # FTP两个业务之间间隔
    "ftp_to_next_sec": 15,        # FTP组完成到下一组间隔
}

# ==================== 置信度参数 ====================
CONFIDENCE_WEIGHTS = {
    "flow_match": 0.4,       # 流量匹配度权重
    "duration_match": 0.3,   # 时长接近度权重
    "order_consistency": 0.2, # 顺序一致性权重
    "group_relation": 0.1,   # 组关系权重
}

CONFIDENCE_THRESHOLD = 0.6   # 置信度阈值

# ==================== GUI 配置 ====================
GUI_CONFIG = {
    "window_width": 1024,
    "window_height": 768,
    "min_width": 800,
    "min_height": 600,
    "title": f"{APP_NAME} V{APP_VERSION}",
}

# ==================== 配色方案 ====================
COLORS = {
    "background": "#F5F5F5",
    "card_background": "#FFFFFF",
    "primary": "#1E88E5",      # 主色调（蓝色）
    "secondary": "#43A047",    # 次色调（绿色）
    "warning": "#FFA726",      # 警告色（橙色）
    "error": "#E53935",        # 错误色（红色）
    "text_primary": "#212121",
    "text_secondary": "#757575",
    "table_alternate": "#FAFAFA",
}

# ==================== 业务颜色 ====================
BUSINESS_COLORS = {
    "ftp_download": "#1E88E5",   # 蓝色
    "ftp_upload": "#42A5F5",     # 浅蓝
    "store_small": "#66BB6A",    # 浅绿
    "store_large": "#43A047",    # 绿色
    "wechat_small": "#FFA726",   # 橙色
    "wechat_large": "#FF7043",   # 深橙
    "pending": "#9E9E9E",        # 灰色（待确认）
}

# ==================== 运营商配置 ====================
OPERATOR_KEYWORDS = {
    "联通": ["联通", "unicom", "lt"],
    "电信": ["电信", "telecom", "ct", "dx"],
}

# ==================== 输出配置 ====================
OUTPUT_CONFIG = {
    "default_filename": "5G速率统计_{timestamp}.xlsx",
    "timestamp_format": "%Y%m%d-%H%M",
    "sheets": {
        "unicom": "联通",
        "telecom": "电信",
        "summary": "汇总",
        "pending": "待确认",
        "version": "版本信息",
    },
}

# ==================== 版本历史 ====================
VERSION_HISTORY = [
    {
        "version": "2.0.0",
        "date": "2026-07-10",
        "changes": [
            "从零全新重写",
            "三页向导GUI设计",
            "六业务自动场景推断",
            "置信度评分与待确认机制",
            "多文件多运营商支持",
            "不依赖基准文件运行",
        ]
    }
]

# ==================== 数据列名映射 ====================
COLUMN_MAPPING = {
    # 时间列
    "time": ["Millisecond", "采集时间", "Time"],

    # 速率列
    "dl_mac": ["Downlink MAC Throughput(bps)", "下行MAC吞吐率(bps)"],
    "ul_mac": ["Uplink MAC Throughput(bps)", "上行MAC吞吐率(bps)"],
    "dl_rlc": ["Downlink RLC Throughput(bps)", "下行RLC吞吐率(bps)"],
    "ul_rlc": ["Uplink RLC Throughput(bps)", "上行RLC吞吐率(bps)"],

    # 流量列
    "dl_flow": ["Downlink Flow(MB)", "下行流量(MB)"],
    "ul_flow": ["Uplink Flow(MB)", "上行流量(MB)"],

    # 信号列
    "rsrp": ["SSB Beam Rsrp", "SRS RSRP", "RSRP"],
    "sinr": ["SINR"],

    # 其他
    "crnti": ["CRNTI"],
    "ca": ["CA"],
}

def get_version_string():
    """获取版本字符串"""
    return f"V{APP_VERSION}"

def get_window_title():
    """获取窗口标题"""
    return f"{APP_NAME} V{APP_VERSION}"

def get_business_name(business_key):
    """获取业务中文名称"""
    return BUSINESS_SPECS.get(business_key, {}).get("name", business_key)

def get_business_direction(business_key):
    """获取业务方向"""
    return BUSINESS_SPECS.get(business_key, {}).get("direction", "未知")

def is_download_business(business_key):
    """判断是否为下载业务"""
    return get_business_direction(business_key) == "下行"

def is_upload_business(business_key):
    """判断是否为上传业务"""
    return get_business_direction(business_key) == "上行"