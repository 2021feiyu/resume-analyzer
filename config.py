# -*- coding: utf-8 -*-
"""
简历评估工具 V2 — 配置文件
============================
工程化版本：异步并发、断点续传、Word支持、SQLite状态管理
"""

# ============================================================
# API 配置
# ============================================================
import os

DASHSCOPE_API_KEY = os.environ.get("DASHSCOPE_API_KEY", "")
DASHSCOPE_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL = "qwen3.7-plus"

# ============================================================
# 路径配置
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RESUME_FOLDER = os.environ.get("RESUME_FOLDER", os.path.join(BASE_DIR, "待评估简历"))
PROMPT_FILE = os.environ.get("PROMPT_FILE", os.path.join(BASE_DIR, "评估提示词模板.txt"))
OUTPUT_EXCEL = os.environ.get("OUTPUT_EXCEL", os.path.join(BASE_DIR, "简历评估结果.xlsx"))
STATE_DB = os.environ.get("STATE_DB", os.path.join(BASE_DIR, "状态.db"))
BACKUP_DIR = os.environ.get("BACKUP_DIR", os.path.join(BASE_DIR, "结果备份"))
ERROR_LOG = os.environ.get("ERROR_LOG", os.path.join(BASE_DIR, "错误日志.log"))
APP_LOG = os.environ.get("APP_LOG", os.path.join(BASE_DIR, "运行日志.log"))

# ============================================================
# 并发控制 (基于 qwen3.7-plus 官方限流: 30,000 RPM / 5,000,000 TPM)
# ============================================================
MAX_CONCURRENT = 5           # 最大并发数 (保守值, 可上调至 10-15)
API_RETRY_COUNT = 3           # API 失败重试次数
API_RETRY_BASE_DELAY = 2.0    # 重试基础延迟 (秒), 实际=base * 2^attempt
API_TIMEOUT = 300             # 单次 API 超时 (秒) — 多模态较慢

# ============================================================
# 图片渲染配置
# ============================================================
PDF_DPI = 150                 # PDF 渲染分辨率
PDF_MAX_PAGES = 4             # 最多渲染页数
PDF_JPEG_QUALITY = 85         # JPEG 质量

# Word 文档渲染
DOCX_FONT_SIZE = 16           # 渲染字号
DOCX_LINE_SPACING = 24        # 行距 (像素)
DOCX_PAGE_WIDTH = 1200        # 渲染画布宽度
DOCX_MARGIN = 40              # 边距 (像素)

# ============================================================
# 文件大小限制
# ============================================================
MAX_FILE_SIZE_MB = 20          # 单文件大小上限 (放宽到 20MB)
