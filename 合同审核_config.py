# -*- coding: utf-8 -*-
"""
合同审核工具 V1 — 配置文件
============================
基于简历评估工具V2改造：异步并发、断点续传、PDF支持、SQLite状态管理
"""

# ============================================================
# API 配置
# ============================================================
import os

DASHSCOPE_API_KEY = os.environ.get("DASHSCOPE_API_KEY", "")
DASHSCOPE_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL = "qwen3.7-plus"

# ============================================================
# 路径配置 (全部指向本项目目录)
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONTRACT_FOLDER = os.environ.get("CONTRACT_FOLDER", os.path.join(BASE_DIR, "待审核合同"))
PROMPT_FILE = os.environ.get("CONTRACT_PROMPT_FILE", os.path.join(BASE_DIR, "合同审核提示词模板.txt"))
OUTPUT_EXCEL = os.environ.get("CONTRACT_OUTPUT_EXCEL", os.path.join(BASE_DIR, "合同审核结果.xlsx"))
STATE_DB = os.environ.get("CONTRACT_STATE_DB", os.path.join(BASE_DIR, "合同状态.db"))
BACKUP_DIR = os.environ.get("CONTRACT_BACKUP_DIR", os.path.join(BASE_DIR, "结果备份"))
ERROR_LOG = os.environ.get("CONTRACT_ERROR_LOG", os.path.join(BASE_DIR, "合同错误日志.log"))
APP_LOG = os.environ.get("CONTRACT_APP_LOG", os.path.join(BASE_DIR, "合同运行日志.log"))

# ============================================================
# 并发控制 (基于 qwen3.7-plus 官方限流: 30,000 RPM / 5,000,000 TPM)
# ============================================================
MAX_CONCURRENT = 5           # 最大并发数
API_RETRY_COUNT = 3           # API 失败重试次数
API_RETRY_BASE_DELAY = 2.0    # 重试基础延迟 (秒), 实际=base * 2^attempt
API_TIMEOUT = 300             # 单次 API 超时 (秒) — 合同页数多，多模态较慢

# ============================================================
# 图片渲染配置
# ============================================================
PDF_DPI = 150                 # PDF 渲染分辨率
PDF_MAX_PAGES = 20            # 最多渲染页数 (劳动合同通常10-15页)
PDF_JPEG_QUALITY = 85         # JPEG 质量

# Word 文档渲染
DOCX_FONT_SIZE = 16           # 渲染字号
DOCX_LINE_SPACING = 24        # 行距 (像素)
DOCX_PAGE_WIDTH = 1200        # 渲染画布宽度
DOCX_MARGIN = 40              # 边距 (像素)

# ============================================================
# 文件大小限制
# ============================================================
MAX_FILE_SIZE_MB = 30          # 单文件大小上限 (合同PDF通常较大)
