# -*- coding: utf-8 -*-
"""
合同审核工具 V1 — 主编排器 (纯数据提取)
============================
异步并发 · 断点续传 · PDF支持 · SQLite状态 · 条件格式Excel

用法: python 合同审核.py

结构:
  配置_v2.py        ← 用户编辑的配置文件
  转换器.py         ← 文档转图片 (PDF/DOCX/DOC)
  合同审核提示词模板.txt ← 劳动合同结构化数据提取提示词
  本文件            ← 主编排 + 状态管理 + API调用 + Excel输出
"""

import re
import sys
import json
import time
import asyncio
import sqlite3
import logging
from pathlib import Path
from datetime import datetime
from typing import Any

import aiohttp
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
try:
    from 配置_v2 import (  # noqa: E402
        DASHSCOPE_API_KEY, DASHSCOPE_BASE_URL, MODEL,
        CONTRACT_FOLDER, PROMPT_FILE, OUTPUT_EXCEL, STATE_DB,
        ERROR_LOG, APP_LOG, BACKUP_DIR,
        MAX_CONCURRENT, API_RETRY_COUNT, API_RETRY_BASE_DELAY, API_TIMEOUT,
        PDF_DPI, PDF_MAX_PAGES, PDF_JPEG_QUALITY,
        DOCX_FONT_SIZE, DOCX_LINE_SPACING, DOCX_PAGE_WIDTH, DOCX_MARGIN,
        MAX_FILE_SIZE_MB,
    )
except ImportError:
    import os as _os
    _BASE = _os.path.dirname(_os.path.abspath(__file__))
    DASHSCOPE_API_KEY    = "sk-your-api-key-here"
    DASHSCOPE_BASE_URL   = "https://dashscope.aliyuncs.com/compatible-mode/v1"
    MODEL                = "qwen3.7-plus"
    CONTRACT_FOLDER      = _os.path.join(_BASE, "待审核合同")
    PROMPT_FILE          = _os.path.join(_BASE, "合同审核提示词模板.txt")
    OUTPUT_EXCEL         = _os.path.join(_BASE, "合同审核结果.xlsx")
    STATE_DB             = _os.path.join(_BASE, "状态.db")
    ERROR_LOG            = _os.path.join(_BASE, "错误日志.log")
    APP_LOG              = _os.path.join(_BASE, "运行日志.log")
    BACKUP_DIR           = _os.path.join(_BASE, "结果备份")
    MAX_CONCURRENT       = 5
    API_RETRY_COUNT      = 3
    API_RETRY_BASE_DELAY = 2.0
    API_TIMEOUT          = 300
    PDF_DPI              = 150
    PDF_MAX_PAGES        = 20
    PDF_JPEG_QUALITY     = 85
    DOCX_FONT_SIZE       = 16
    DOCX_LINE_SPACING    = 24
    DOCX_PAGE_WIDTH      = 1200
    DOCX_MARGIN          = 40
    MAX_FILE_SIZE_MB     = 30

# ---------------------------------------------------------------------------
# 日志
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(APP_LOG, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)

err_fh = logging.FileHandler(ERROR_LOG, encoding="utf-8")
err_fh.setLevel(logging.ERROR)
err_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(err_fh)

# 外部模块
from 转换器 import convert_to_images, file_hash  # noqa: E402


# ============================================================
# SQLite 状态管理器
# ============================================================

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS contract_state (
    file_hash   TEXT PRIMARY KEY,
    file_name   TEXT NOT NULL,
    file_path   TEXT NOT NULL,
    status      TEXT NOT NULL DEFAULT 'pending',
    risk_score  REAL,
    risk_level  TEXT,
    party_a     TEXT,
    party_b     TEXT,
    content     TEXT,
    tokens_in   INTEGER,
    tokens_out  INTEGER,
    elapsed_s   REAL,
    error_msg   TEXT,
    created_at  TEXT NOT NULL,
    updated_at  TEXT NOT NULL
)
"""


class StateManager:
    """SQLite 断点续传状态管理。"""

    def __init__(self, db_path: str):
        self.db_path = Path(db_path)
        self._init_db()

    def _init_db(self) -> None:
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            conn.execute(CREATE_TABLE_SQL)
            conn.execute("CREATE INDEX IF NOT EXISTS idx_status ON contract_state(status)")
            cols = [r[1] for r in conn.execute("PRAGMA table_info(contract_state)")]
            if "content" not in cols:
                conn.execute("ALTER TABLE contract_state ADD COLUMN content TEXT")
            if "risk_score" not in cols:
                conn.execute("ALTER TABLE contract_state ADD COLUMN risk_score REAL")
            if "risk_level" not in cols:
                conn.execute("ALTER TABLE contract_state ADD COLUMN risk_level TEXT")
            if "party_a" not in cols:
                conn.execute("ALTER TABLE contract_state ADD COLUMN party_a TEXT")
            if "party_b" not in cols:
                conn.execute("ALTER TABLE contract_state ADD COLUMN party_b TEXT")
            conn.commit()

    def is_completed(self, file_hash: str) -> bool:
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            r = conn.execute(
                "SELECT 1 FROM contract_state WHERE file_hash=? AND status='completed'",
                (file_hash,),
            ).fetchone()
            return r is not None

    def get_stats(self) -> dict:
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            total = conn.execute("SELECT COUNT(*) FROM contract_state").fetchone()[0]
            done = conn.execute("SELECT COUNT(*) FROM contract_state WHERE status='completed'").fetchone()[0]
            failed = conn.execute("SELECT COUNT(*) FROM contract_state WHERE status='failed'").fetchone()[0]
            skipped = conn.execute("SELECT COUNT(*) FROM contract_state WHERE status='skipped'").fetchone()[0]
            avg = conn.execute("SELECT AVG(elapsed_s) FROM contract_state WHERE status='completed'").fetchone()[0] or 0
            return {"total": total, "completed": done, "failed": failed,
                    "pending": total - done - failed - skipped, "avg_elapsed_s": avg}

    def mark_pending(self, file_hash: str, file_name: str, file_path: str) -> None:
        """注册文件为 pending。已存在则重置状态（允许重试之前失败的文件）。"""
        now = datetime.now().isoformat()
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            conn.execute(
                "INSERT INTO contract_state (file_hash,file_name,file_path,status,created_at,updated_at) "
                "VALUES (?,?,?,'pending',?,?) "
                "ON CONFLICT(file_hash) DO UPDATE SET status='pending',error_msg=NULL,updated_at=?",
                (file_hash, file_name, file_path, now, now, now),
            )
            conn.commit()

    def mark_completed(self, file_hash: str, risk_score: float = 0, risk_level: str = "",
                       party_a: str = "", party_b: str = "", content: str = "",
                       tokens_in: int = 0, tokens_out: int = 0, elapsed_s: float = 0) -> None:
        now = datetime.now().isoformat()
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            conn.execute(
                "UPDATE contract_state SET status='completed',risk_score=?,risk_level=?,party_a=?,party_b=?,content=?,tokens_in=?,tokens_out=?,elapsed_s=?,updated_at=? WHERE file_hash=?",
                (risk_score, risk_level, party_a, party_b, content, tokens_in, tokens_out, elapsed_s, now, file_hash),
            )
            conn.commit()

    def mark_failed(self, file_hash: str, error_msg: str = "") -> None:
        now = datetime.now().isoformat()
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            conn.execute("UPDATE contract_state SET status='failed',error_msg=?,updated_at=? WHERE file_hash=?",
                         (error_msg[:500], now, file_hash))
            conn.commit()

    def mark_skipped(self, file_hash: str, reason: str = "") -> None:
        now = datetime.now().isoformat()
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            conn.execute("UPDATE contract_state SET status='skipped',error_msg=?,updated_at=? WHERE file_hash=?",
                         (reason, now, file_hash))
            conn.commit()

    def get_all_results(self) -> list[dict]:
        with sqlite3.connect(str(self.db_path), timeout=30) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(
                "SELECT * FROM contract_state WHERE status IN ('completed','failed') ORDER BY file_name"
            ).fetchall()
            return [dict(r) for r in rows]


# ============================================================
# 异步评估器
# ============================================================

class AsyncEvaluator:
    """asyncio 并发 qwen3.7-plus 多模态合同审核器。"""

    def __init__(self, api_key: str, base_url: str, model: str = "qwen3.7-plus",
                 max_concurrent: int = 5, retry_count: int = 3,
                 retry_base_delay: float = 2.0, timeout: int = 300):
        self.api_key = api_key
        self.base_url = base_url.rstrip("/") + "/chat/completions"
        self.model = model
        self.semaphore = asyncio.Semaphore(max_concurrent)
        self.retry_count = retry_count
        self.retry_base_delay = retry_base_delay
        self.timeout = aiohttp.ClientTimeout(total=timeout)
        self._session: aiohttp.ClientSession | None = None

    async def _get_session(self) -> aiohttp.ClientSession:
        """复用 HTTP session，避免每次请求新建连接池。"""
        if self._session is None or self._session.closed:
            self._session = aiohttp.ClientSession(timeout=self.timeout)
        return self._session

    async def close(self) -> None:
        """关闭 session，释放连接池资源。"""
        if self._session and not self._session.closed:
            await self._session.close()

    async def evaluate(self, page_images: list[str], prompt_template: str,
                       file_name: str) -> dict[str, Any]:
        async with self.semaphore:
            return await self._retry(page_images, prompt_template, file_name)

    async def _retry(self, page_images: list[str], prompt_template: str,
                     file_name: str) -> dict[str, Any]:
        last_error = ""
        for attempt in range(1 + self.retry_count):
            try:
                return await self._call(page_images, prompt_template, file_name)
            except asyncio.TimeoutError:
                last_error = "API timeout"
            except aiohttp.ClientError as e:
                last_error = f"HTTP: {e}"
            except Exception as e:
                last_error = f"{type(e).__name__}: {e}"
            if attempt < self.retry_count:
                logger.warning(f"  [重试 {attempt+1}/{self.retry_count}] {file_name}: {last_error}")
                await asyncio.sleep(self.retry_base_delay * (2 ** attempt))
        return {"status": "failed", "content": "", "tokens_in": 0,
                "tokens_out": 0, "elapsed_s": 0, "error": last_error}

    async def _call(self, page_images: list[str], prompt_template: str,
                    file_name: str) -> dict[str, Any]:
        user_content: list[dict] = []
        for url in page_images:
            user_content.append({"type": "image_url", "image_url": {"url": url}})
        user_content.append({
            "type": "text",
            "text": f"请严格按照上述JSON格式，从这份合同（文件名：{file_name}，共 {len(page_images)} 页）中先识别合同类型，再精确提取对应的结构化数据。逐页阅读、逐字段提取，只输出JSON。",
        })

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": prompt_template},
                {"role": "user", "content": user_content},
            ],
            "temperature": 0.1,
            "max_tokens": 8192,
        }
        headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}

        t0 = time.monotonic()
        session = await self._get_session()
        try:
            async with session.post(self.base_url, json=payload, headers=headers) as resp:
                elapsed_s = round(time.monotonic() - t0, 1)
                if resp.status == 200:
                    data = await resp.json()
                    choices = data.get("choices", [])
                    if not choices:
                        raise aiohttp.ClientError("API returned empty choices")
                    choice = choices[0]
                    message = choice.get("message", {})
                    content = message.get("content", "")
                    if not content:
                        raise aiohttp.ClientError("API returned empty content")
                    usage = data.get("usage", {})
                    logger.info(f"  [完成] {file_name} | in={usage.get('prompt_tokens',0)} out={usage.get('completion_tokens',0)} ({elapsed_s}s)")
                    return {"status": "completed", "content": content,
                            "tokens_in": usage.get("prompt_tokens", 0),
                            "tokens_out": usage.get("completion_tokens", 0),
                            "elapsed_s": elapsed_s, "error": ""}
                else:
                    err = await resp.text()
                    raise aiohttp.ClientError(f"HTTP {resp.status}: {err[:200]}")
        except Exception:
            # 响应异常时关闭当前 session，下次请求会自动新建
            if self._session and not self._session.closed:
                await self._session.close()
                self._session = None
            raise


# ============================================================
# Excel 输出
# ============================================================

def _parse_json_safe(content: str) -> dict:
    """从模型响应中安全提取 JSON。"""
    if not content:
        return {}
    try:
        return json.loads(content)
    except Exception:
        for pat in [r"```(?:json)?\s*\n?(.*?)\n?```", r"\{.*\}"]:
            m = re.search(pat, content, re.DOTALL)
            if m:
                try:
                    return json.loads(m.group(1) if "```" in pat else m.group(0))
                except Exception:
                    continue
    return {}


# 字段映射: JSON key → Excel 列名（44个提取字段）
FIELD_MAP = [
    # -- 合同载体 --
    ("合同载体", "合同载体"),
    # -- 合同类型识别 --
    ("合同类型", "合同类型"),
    # -- 甲方 --
    ("甲方全称", "甲方全称"),
    ("法定代表人", "法定代表人"),
    # -- 乙方 --
    ("乙方姓名", "乙方姓名"),
    ("身份证号", "身份证号"),
    ("联系电话", "联系电话"),
    ("电子邮件", "电子邮件"),
    # -- 协议基本信息 --
    ("合同起始日期", "合同起始日期"),
    ("合同终止日期", "合同终止日期"),
    ("岗位名称", "岗位名称"),
    ("主要工作地点", "主要工作地点"),
    # -- 试用期 --
    ("试用期起始日期", "试用期起始日期"),
    ("试用期终止日期", "试用期终止日期"),
    ("试用期长度", "试用期长度"),
    # -- 薪酬（标准劳动合同） --
    ("转正月总薪酬", "转正月总薪酬"),
    ("转正基本工资", "转正基本工资"),
    ("转正绩效工资", "转正绩效工资"),
    ("转正驻场补贴", "转正驻场补贴"),
    ("试用期月总薪酬", "试用期月总薪酬"),
    ("试用期基本工资", "试用期基本工资"),
    ("试用期绩效工资", "试用期绩效工资"),
    ("试用期驻场补贴", "试用期驻场补贴"),
    ("试用期薪酬占比", "试用期薪酬占比"),
    ("转正基本工资占比", "转正基本工资占比"),
    ("试用期基本工资占比", "试用期基本工资占比"),
    ("其他补贴", "其他补贴"),
    ("是否涉及派驻", "是否涉及派驻"),
    ("缴纳地", "缴纳地"),
    # -- 实习协议专属 --
    ("实习基本津贴", "实习基本津贴"),
    ("实习绩效津贴", "实习绩效津贴"),
    ("协议关系说明", "协议关系说明"),
    # -- 日薪制专属 --
    ("日薪标准", "日薪标准"),
    ("月计薪天数", "月计薪天数"),
    ("试用期日薪", "试用期日薪"),
    ("发薪周期", "发薪周期"),
    # -- 计件制专属 --
    ("计件单价", "计件单价"),
    ("计件单位", "计件单位"),
    ("月保底工资", "月保底工资"),
    ("计件绩效说明", "计件绩效说明"),
    ("试用期计件单价", "试用期计件单价"),
    # -- 签署 --
    ("合同签订日期", "合同签订日期"),
    ("合同版本号", "合同版本号"),
]


def _truncate(val, max_len: int = 50) -> str:
    """截断过长的字符串值。"""
    s = str(val or "")
    return s[:max_len] if len(s) > max_len else s


def write_excel(records: list[dict], output_path: str) -> None:
    """写入 Excel，44个信息提取字段（4种合同类型）。"""
    if not records:
        logger.warning("没有记录可写入 Excel")
        return

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    rows = []
    STATUS_CN = {"completed": "完成", "failed": "失败", "pending": "待处理", "skipped": "已跳过"}
    for r in records:
        parsed = _parse_json_safe(r.get("content", ""))
        raw_status = str(r.get("status", ""))
        row = {
            "文件名": r.get("file_name", ""),
            "状态": STATUS_CN.get(raw_status, raw_status),
        }
        for json_key, col_name in FIELD_MAP:
            row[col_name] = _truncate(parsed.get(json_key, ""))
        error_msg = str(r.get("error_msg", ""))
        row["错误信息"] = error_msg[:100] if error_msg else ""
        row["处理时间"] = str(r.get("updated_at", ""))[:19].replace("T", " ")
        rows.append(row)

    df = pd.DataFrame(rows)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="合同提取结果")
        ws = writer.sheets["合同提取结果"]

        # 列宽
        for col_idx in range(1, len(df.columns) + 1):
            max_len = 0
            for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
                for cv in row:
                    if cv:
                        max_len = max(max_len, sum(2 if ord(c) > 127 else 1 for c in str(cv)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 50)

        ws.freeze_panes = "A2"

        # 表头
        hf = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        hfont = Font(color="FFFFFF", bold=True, size=11)
        for cell in ws[1]:
            cell.fill = hf
            cell.font = hfont
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # 失败行标红
        status_col = None
        for i, name in enumerate(df.columns, 1):
            if name == "状态":
                status_col = i
                break
        if status_col:
            for r in range(2, len(df) + 2):
                if str(ws.cell(row=r, column=status_col).value) == "失败":
                    for c in range(1, len(df.columns) + 1):
                        ws.cell(row=r, column=c).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    logger.info(f"Excel 已保存: {out} ({len(df)} 行)")


# ============================================================
# 主流程
# ============================================================

def validate(check_api_key: bool = True) -> None:
    if check_api_key and DASHSCOPE_API_KEY == "sk-your-api-key-here":
        logger.error("请在 配置_v2.py 中设置 DASHSCOPE_API_KEY!")
        sys.exit(1)
    for p, n in [(PROMPT_FILE, "Prompt"), (CONTRACT_FOLDER, "合同文件夹")]:
        if not Path(p).exists():
            logger.error(f"{n} 不存在: {p}")
            sys.exit(1)


def scan_and_register(state: StateManager) -> list[dict]:
    folder = Path(CONTRACT_FOLDER)
    max_bytes = MAX_FILE_SIZE_MB * 1024 * 1024
    pending: list[dict] = []
    skipped_large = skipped_done = 0

    for ext in ("*.pdf", "*.docx", "*.doc"):
        for f in sorted(folder.glob(ext)):
            try:
                h = file_hash(f)
            except Exception as e:
                logger.warning(f"  跳过(无法读取): {f.name} — {e}")
                continue
            try:
                size_mb = f.stat().st_size / (1024 * 1024)
            except Exception as e:
                logger.warning(f"  跳过(无法获取大小): {f.name} — {e}")
                continue
            if f.stat().st_size > max_bytes:
                state.mark_skipped(h, f"过大 ({size_mb:.1f}MB)")
                skipped_large += 1
                continue
            if state.is_completed(h):
                skipped_done += 1
                continue
            state.mark_pending(h, f.name, str(f))
            pending.append({"hash": h, "name": f.name, "path": f})

    stats = state.get_stats()
    logger.info(f"扫描: 新增 {len(pending)} | 已完成 {skipped_done} | 跳过 {skipped_large} | 总计 {stats['total']}")
    return pending


async def process_one(evaluator: AsyncEvaluator, state: StateManager,
                      item: dict, prompt: str) -> None:
    fhash, fname, fpath = item["hash"], item["name"], item["path"]
    try:
        images = convert_to_images(fpath, dpi=PDF_DPI, max_pages=PDF_MAX_PAGES,
                                   quality=PDF_JPEG_QUALITY, page_width=DOCX_PAGE_WIDTH,
                                   margin=DOCX_MARGIN, font_size=DOCX_FONT_SIZE,
                                   line_spacing=DOCX_LINE_SPACING)
        if not images:
            state.mark_failed(fhash, "文档转换失败")
            return

        result = await evaluator.evaluate(images, prompt, fname)
        if result["status"] == "completed":
            content = result["content"] or ""
            parsed = _parse_json_safe(content)
            state.mark_completed(
                fhash,
                risk_score=0,
                risk_level="",
                party_a=str(parsed.get("甲方全称", "")),
                party_b=str(parsed.get("乙方姓名", "")),
                content=content,
                tokens_in=result["tokens_in"],
                tokens_out=result["tokens_out"],
                elapsed_s=result["elapsed_s"],
            )
            logger.info(f"  [OK] {fname} -> {parsed.get('乙方姓名','')} | "
                        f"{parsed.get('岗位名称','')} | "
                        f"转正{parsed.get('转正月总薪酬','')}元 | "
                        f"{result['elapsed_s']}s")
        else:
            state.mark_failed(fhash, result.get("error", ""))
            logger.error(f"  [FAIL] {fname}: {result.get('error','')}")
    except Exception as e:
        state.mark_failed(fhash, f"{type(e).__name__}: {e}")
        logger.error(f"  [FAIL] {fname}: {e}")


async def main_async(export_only: bool = False) -> None:
    logger.info("=" * 60)
    logger.info(f"合同信息提取 V1 | {MODEL} | 并发:{MAX_CONCURRENT}")
    logger.info(f"目录: {CONTRACT_FOLDER}")
    logger.info("=" * 60)

    state = StateManager(STATE_DB)

    # --export: 仅从 DB 导出，不调用 API
    if export_only:
        logger.info("模式: 从 DB 全量导出 Excel")
        validate(check_api_key=False)
        Path(BACKUP_DIR).mkdir(parents=True, exist_ok=True)
        export_path = str(Path(BACKUP_DIR) / f"合同审核结果_全量导出_{datetime.now():%Y%m%d_%H%M}.xlsx")
        records = state.get_all_results()
        if records:
            write_excel(records, export_path)
        else:
            logger.warning("DB 中无记录")
        return

    validate()
    evaluator = AsyncEvaluator(DASHSCOPE_API_KEY, DASHSCOPE_BASE_URL, MODEL,
                               MAX_CONCURRENT, API_RETRY_COUNT, API_RETRY_BASE_DELAY, API_TIMEOUT)

    prompt = Path(PROMPT_FILE).read_text(encoding="utf-8")
    logger.info(f"Prompt: {len(prompt)} 字符")

    pending = scan_and_register(state)
    if not pending:
        logger.info("全部完成!")
        _finalize(state)
        return

    total = len(pending)
    logger.info(f"开始处理 {total} 份合同\n")
    t_start = time.monotonic()

    for i in range(0, total, MAX_CONCURRENT):
        batch = pending[i:i + MAX_CONCURRENT]
        logger.info(f"--- 批次 {i // MAX_CONCURRENT + 1} ({len(batch)} 个) ---")
        await asyncio.gather(*[process_one(evaluator, state, item, prompt) for item in batch])
        if i + MAX_CONCURRENT < total:
            await asyncio.sleep(1.0)

    logger.info(f"\n总耗时: {time.monotonic() - t_start:.0f}s")
    _finalize(state)
    await evaluator.close()


def _finalize(state: StateManager) -> None:
    stats = state.get_stats()
    logger.info(f"统计: 完成 {stats['completed']} | 失败 {stats['failed']} | 总计 {stats['total']}")
    if stats["completed"] > 0:
        logger.info(f"平均: {stats['avg_elapsed_s']:.1f}s/份")
    records = state.get_all_results()
    if records:
        Path(BACKUP_DIR).mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        export_path = str(Path(BACKUP_DIR) / f"合同审核结果_{ts}.xlsx")
        write_excel(records, export_path)
        logger.info(f"结果已备份: {export_path}")


def main() -> None:
    export_only = "--export" in sys.argv
    asyncio.run(main_async(export_only=export_only))


if __name__ == "__main__":
    main()
