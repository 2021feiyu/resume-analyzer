# -*- coding: utf-8 -*-
"""
简历评估工具 V2 — 文档转换器
==============================
PDF  → PyMuPDF 渲染为 JPEG
DOCX → python-docx 提取文本 → PIL 渲染为 PNG
DOC  → olefile 解析 OLE2 → PIL 渲染为 PNG
统一输出: list[str] (base64 data URL)
"""

import re
import base64
import hashlib
import logging
from io import BytesIO
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

logger = logging.getLogger(__name__)

# ============================================================
# 文件哈希 (用于断点续传去重)
# ============================================================

def file_hash(path: Path) -> str:
    """计算文件 SHA256 哈希。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


# ============================================================
# 字体 & 渲染工具 (DOCX/DOC 共用)
# ============================================================

def _get_font(size: int = 16) -> ImageFont.FreeTypeFont:
    """获取中文字体 (Windows 微软雅黑)。"""
    font_paths = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simsun.ttc",
        "C:/Windows/Fonts/simhei.ttf",
    ]
    for fp in font_paths:
        if Path(fp).exists():
            return ImageFont.truetype(fp, size)
    return ImageFont.load_default()


def _split_text_lines(text: str, font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    """按像素宽度自动换行。"""
    lines = []
    for paragraph in text.split("\n"):
        if not paragraph.strip():
            lines.append("")
            continue
        current = ""
        for char in paragraph:
            test = current + char
            bbox = font.getbbox(test)
            if bbox[2] > max_width:
                lines.append(current)
                current = char
            else:
                current = test
        if current:
            lines.append(current)
    return lines


def _render_text_to_images(
    all_lines: list[str],
    page_width: int = 1200,
    margin: int = 40,
    font_size: int = 16,
    line_spacing: int = 24,
) -> list[str]:
    """将文本行列表渲染为 base64 PNG 图片 (自动分页)。"""
    font = _get_font(font_size)
    usable_height = page_width * 1.414 - 2 * margin  # A4 比例
    lines_per_page = max(1, int((usable_height - line_spacing) / line_spacing))

    images: list[str] = []
    for page_start in range(0, len(all_lines), lines_per_page):
        page_lines = all_lines[page_start:page_start + lines_per_page]

        img = Image.new("RGB", (page_width, int(page_width * 1.414)), "white")
        draw = ImageDraw.Draw(img)

        y = margin
        for line in page_lines:
            if line:
                draw.text((margin, y), line, fill="black", font=font)
            y += line_spacing

        buf = BytesIO()
        img.save(buf, format="PNG", optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        images.append(f"data:image/png;base64,{b64}")

    return images


# ============================================================
# PDF → 图片
# ============================================================

def pdf_to_images(
    pdf_path: Path,
    dpi: int = 150,
    max_pages: int = 4,
    quality: int = 85,
) -> list[str]:
    """将 PDF 逐页渲染为 base64 JPEG。"""
    doc = fitz.open(pdf_path)
    try:
        total_pages = len(doc)
        pages = min(total_pages, max_pages)
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        images: list[str] = []
        for i in range(pages):
            pix = doc[i].get_pixmap(matrix=mat)
            img_bytes = pix.tobytes("jpeg", jpg_quality=quality)
            b64 = base64.b64encode(img_bytes).decode("ascii")
            images.append(f"data:image/jpeg;base64,{b64}")
        logger.debug(f"  PDF→图片: {pdf_path.name} {total_pages}页→{len(images)}张")
        return images
    finally:
        doc.close()


# ============================================================
# DOCX → 图片
# ============================================================

def docx_to_images(
    docx_path: Path,
    page_width: int = 1200,
    margin: int = 40,
    font_size: int = 16,
    line_spacing: int = 24,
) -> list[str]:
    """将 DOCX 文本提取后渲染为 PNG 图片。"""
    try:
        from docx import Document
    except ImportError:
        raise ImportError("请安装 python-docx: pip install python-docx")

    doc = Document(docx_path)
    font = _get_font(font_size)
    text_width = page_width - 2 * margin

    all_lines: list[str] = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            all_lines.append("")
        else:
            wrapped = _split_text_lines(text, font, text_width)
            all_lines.extend(wrapped)
        all_lines.append("")  # 段落间距

    if not all_lines:
        all_lines = ["(空白文档)"]

    images = _render_text_to_images(all_lines, page_width, margin, font_size, line_spacing)
    logger.debug(f"  DOCX→图片: {docx_path.name} {len(all_lines)}行→{len(images)}张")
    return images


# ============================================================
# DOC (旧版 Word) → 图片
# ============================================================

def _extract_doc_text(doc_path: Path) -> str:
    """从 .doc (OLE2 格式) 文件中提取纯文本。

    先尝试用 python-docx (部分 .doc 实际是 .docx),
    失败后用 olefile 解析 OLE2 复合文档提取 WordDocument 流。
    """
    # 方案 1: 尝试 python-docx (部分 .doc 是伪装成 .doc 的 .docx)
    try:
        from docx import Document
        doc = Document(doc_path)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        pass

    # 方案 2: olefile 解析 OLE2 复合文档
    try:
        import olefile
        ole = olefile.OleFileIO(doc_path)
        try:
            if ole.exists("WordDocument"):
                data = ole.openstream("WordDocument").read()
                text_parts = []
                try:
                    decoded = data.decode("utf-16-le", errors="ignore")
                    for chunk in decoded.split("\x00"):
                        chunk = chunk.strip()
                        if len(chunk) > 2 and any(
                            c.isalpha() or '\u4e00' <= c <= '\u9fff'
                            for c in chunk
                        ):
                            text_parts.append(chunk)
                except Exception:
                    pass

                for stream_name in ["1Table", "0Table"]:
                    if ole.exists(stream_name):
                        try:
                            extra = ole.openstream(stream_name).read()
                            ed = extra.decode("utf-16-le", errors="ignore")
                            for chunk in ed.split("\x00"):
                                chunk = chunk.strip()
                                if len(chunk) > 3:
                                    text_parts.append(chunk)
                        except Exception:
                            pass

                if text_parts:
                    seen = set()
                    unique = []
                    for p in text_parts:
                        if p not in seen:
                            seen.add(p)
                            unique.append(p)
                    result = "\n".join(unique)
                    logger.debug(f"  DOC→文本: {doc_path.name} 提取 {len(result)} 字符")
                    return result
        finally:
            ole.close()
    except ImportError:
        logger.warning("  olefile 未安装, 无法解析 .doc 文件")
    except Exception as e:
        logger.debug(f"  olefile 解析失败: {e}")

    # 方案 3: 原始二进制兜底 — 搜索可读文本
    try:
        with open(doc_path, "rb") as f:
            raw = f.read()
        # 尝试 UTF-16LE 解码
        text = raw.decode("utf-16-le", errors="ignore")
        # 过滤可打印段
        parts = re.findall(r"[\u4e00-\u9fff\w\s.,;:!?()（）、。，；：！？《》\-\+/@#$%^&*\[\]{}|~`'\"\\]{4,}", text)
        if parts:
            result = "\n".join(parts[:200])  # 限制提取量
            logger.debug(f"  DOC→文本(原始): {doc_path.name} 提取 {len(result)} 字符")
            return result
    except Exception as e:
        logger.debug(f"  原始提取失败: {e}")

    return ""


def doc_to_images(
    doc_path: Path,
    page_width: int = 1200,
    margin: int = 40,
    font_size: int = 16,
    line_spacing: int = 24,
) -> list[str]:
    """将 .doc 文件提取文本后渲染为 PNG 图片。"""
    text = _extract_doc_text(doc_path)
    if not text:
        raise ValueError(f"无法从 {doc_path.name} 提取文本")

    font = _get_font(font_size)
    text_width = page_width - 2 * margin

    all_lines: list[str] = []
    for paragraph in text.split("\n"):
        stripped = paragraph.strip()
        if not stripped:
            all_lines.append("")
        else:
            wrapped = _split_text_lines(stripped, font, text_width)
            all_lines.extend(wrapped)
        all_lines.append("")

    if not all_lines:
        all_lines = ["(无法提取文本)"]

    images = _render_text_to_images(all_lines, page_width, margin, font_size, line_spacing)
    logger.debug(f"  DOC→图片: {doc_path.name} {len(all_lines)}行→{len(images)}张")
    return images


# ============================================================
# 统一入口
# ============================================================

def convert_to_images(
    file_path: Path,
    dpi: int = 150,
    max_pages: int = 4,
    quality: int = 85,
    **kwargs,
) -> Optional[list[str]]:
    """统一入口：根据文件扩展名自动选择转换器。

    支持格式: .pdf, .docx, .doc

    Args:
        file_path:  文件路径
        dpi:        PDF 渲染 DPI
        max_pages:  最大页数 (仅 PDF)
        quality:    JPEG 质量 (仅 PDF)
        **kwargs:   传给 Word 渲染器 (page_width, margin, font_size, line_spacing)

    Returns:
        base64 data URL 列表, 失败返回 None
    """
    suffix = file_path.suffix.lower()

    try:
        if suffix == ".pdf":
            return pdf_to_images(file_path, dpi=dpi, max_pages=max_pages, quality=quality)
        elif suffix == ".docx":
            return docx_to_images(file_path, **kwargs)
        elif suffix == ".doc":
            return doc_to_images(file_path, **kwargs)
        else:
            logger.warning(f"  不支持的文件格式: {suffix}")
            return None
    except Exception as e:
        logger.error(f"  转换失败 [{file_path.name}]: {type(e).__name__}: {e}")
        return None
