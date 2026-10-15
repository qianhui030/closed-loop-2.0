# -*- coding: utf-8 -*-
"""Shared paths, config, logging, text/time helpers and safe file writing."""
from __future__ import annotations

import json
import logging
import os
import re
import shutil
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml

# CLOSEDLOOP_ROOT lets the self-test run against a throwaway tree instead of live data.
ROOT = Path(os.environ.get("CLOSEDLOOP_ROOT") or Path(__file__).resolve().parent.parent).resolve()

DIRS = {
    "data_combination": ROOT / "Data Combination",
    "receiving_input": ROOT / "Receiving_Log" / "Input",
    "receiving_std": ROOT / "Receiving_Log" / "Standardized",
    "pool_dir": ROOT / "No_Ascan_Pool",
    "pool_backups": ROOT / "No_Ascan_Pool" / "Backups",
    "step1_output": ROOT / "Step1_Output",
    "status_input": ROOT / "Status_Check" / "Input",
    "final_output": ROOT / "Final_Output",
    "archive": ROOT / "Archive",
    "logs": ROOT / "Logs",
    "working": ROOT / "Working",
    "config": ROOT / "config",
}
POOL_FILE = DIRS["pool_dir"] / "No_Ascan_Pool.xlsx"
POOL_ACTIVE_PARQUET = DIRS["pool_dir"] / "Active_Pool.parquet"
POOL_CLOSED_PARQUET = DIRS["pool_dir"] / "Closed_History.parquet"
RECEIVING_HISTORY = DIRS["receiving_std"] / "receiving_history.csv"

POOL_ACTIVE_SHEET = "Active_Pool"
POOL_CLOSED_SHEET = "Closed_History"

EMPTY_TEXT = {"", "nan", "none", "<na>", "nat", "null", "n/a", "#n/a", "-"}


class StopWithMessage(RuntimeError):
    """Business-level stop. Printed as a clear message instead of a traceback."""


def ensure_dirs() -> None:
    for path in DIRS.values():
        path.mkdir(parents=True, exist_ok=True)


def load_config(name: str) -> dict:
    path = DIRS["config"] / name
    if not path.exists():
        raise StopWithMessage(f"缺少配置文件：{path.relative_to(ROOT)}")
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def setup_logging(run_id: str, stage: str) -> logging.Logger:
    DIRS["logs"].mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(f"closedloop.{stage}")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.propagate = False
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S")
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(fmt)
    logger.addHandler(stream)
    handler = logging.FileHandler(DIRS["logs"] / f"{stage}_{run_id}.log", encoding="utf-8")
    handler.setFormatter(fmt)
    logger.addHandler(handler)
    return logger


def new_run_id(stage: str) -> str:
    return f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{stage}"


# ---------------------------------------------------------------- text helpers

def norm_key(value) -> str:
    """Normalized comparison key for header names: NFKC, no spaces, lowercase."""
    text = unicodedata.normalize("NFKC", str(value)).strip().lower()
    return re.sub(r"\s+", "", text)


def clean_text(value) -> str | None:
    """Trimmed text, or None for any flavour of blank. Never returns 'nan'."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if text.lower() in EMPTY_TEXT:
        return None
    return text


def clean_tno(value) -> str | None:
    """TNO stays text. Strip a spurious .0 from Excel, never change case or pad."""
    text = clean_text(value)
    if text is None:
        return None
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    return text


def mawb_match_key(value) -> str | None:
    """Key for MAWB comparison only. The untouched original stays in mawb_raw."""
    text = clean_text(value)
    if text is None:
        return None
    return re.sub(r"[^0-9A-Za-z]", "", text).upper() or None


def join_unique(values, sep: str = ";") -> str | None:
    out: list[str] = []
    for value in values:
        text = clean_text(value)
        if text is None:
            continue
        for piece in text.split(sep):
            piece = piece.strip()
            if piece and piece not in out:
                out.append(piece)
    return sep.join(out) if out else None


def parse_dt(value, formats: list[str] | None = None):
    """Parse a datetime with explicit formats first. Returns NaT when unusable."""
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value)
    text = clean_text(value)
    if text is None:
        return pd.NaT
    for fmt in formats or []:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue
    parsed = pd.to_datetime(text, errors="coerce")
    return parsed if not isinstance(parsed, pd.Series) else pd.NaT


def parse_dt_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    text = series.astype("string").str.strip()
    text = text.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA, "<NA>": pd.NA})
    return pd.to_datetime(text, errors="coerce")


def fmt_dt_series(series: pd.Series, fmt: str = "%Y-%m-%d %H:%M:%S") -> pd.Series:
    dt = parse_dt_series(series)
    return dt.dt.strftime(fmt).astype("object").where(dt.notna(), None)


def has_timezone(value) -> bool:
    text = clean_text(value)
    if text is None:
        return False
    return bool(re.search(r"(?:[+-]\d{2}:?\d{2}|\bZ\b|UTC|GMT)$", text))


# ---------------------------------------------------------------- header lookup

def find_header_row(rows: list[list], aliases: dict, scan_rows: int) -> int:
    """Row index whose cells hit the most configured aliases. Handles two-row headers."""
    wanted = {norm_key(a) for names in aliases.values() for a in names}
    best_row, best_hits = 0, -1
    for idx, row in enumerate(rows[:scan_rows]):
        hits = sum(1 for cell in row if cell is not None and norm_key(cell) in wanted)
        if hits > best_hits:
            best_row, best_hits = idx, hits
    return best_row


def resolve_columns(columns: list[str], aliases: dict) -> dict[str, str]:
    """Map logical name -> actual column, by exact normalized match only."""
    by_norm: dict[str, str] = {}
    for col in columns:
        by_norm.setdefault(norm_key(col), col)
    found = {}
    for logical, names in aliases.items():
        for name in names:
            hit = by_norm.get(norm_key(name))
            if hit is not None:
                found[logical] = hit
                break
    return found


# ---------------------------------------------------------------- safe file I/O

def assert_writable(path: Path) -> None:
    """Fail loudly when Excel holds the file, rather than silently losing data."""
    if not path.exists():
        return
    try:
        with open(path, "ab"):
            pass
    except PermissionError as exc:
        raise StopWithMessage(
            f"文件被占用，无法写入：{path}\n"
            "请关闭 Excel 中打开的这个文件后重新运行。本次未修改任何已有数据。"
        ) from exc


def backup_file(path: Path, backup_dir: Path, keep: int = 30) -> Path | None:
    if not path.exists():
        return None
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = backup_dir / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, target)
    existing = sorted(backup_dir.glob(f"{path.stem}_*{path.suffix}"))
    for old in existing[:-keep] if keep and len(existing) > keep else []:
        old.unlink(missing_ok=True)
    return target


def atomic_write_parquet(path: Path, frame: pd.DataFrame) -> None:
    assert_writable(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.stem}.tmp{path.suffix}")
    if tmp.exists():
        tmp.unlink()
    frame.to_parquet(tmp, index=False)
    os.replace(tmp, path)


def atomic_write_excel(path: Path, sheets: dict[str, pd.DataFrame], verify: bool = True) -> None:
    """Write to a temp file, verify it reopens with the right shape, then replace."""
    assert_writable(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.stem}.tmp{path.suffix}")
    if tmp.exists():
        tmp.unlink()
    total_rows = sum(len(frame) for frame in sheets.values())
    # Re-reading 30万行 Pool with openpyxl can take as long as the write itself.
    if verify and total_rows > 20000:
        verify = False
    with pd.ExcelWriter(tmp, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
            writer.sheets[name[:31]].freeze_panes(1, 0)
    if verify:
        check = pd.read_excel(tmp, sheet_name=None, engine="openpyxl")
        for name, frame in sheets.items():
            got = check.get(name[:31])
            if got is None:
                raise StopWithMessage(f"写入校验失败：{path.name} 缺少工作表 {name}")
            if len(got) != len(frame):
                raise StopWithMessage(
                    f"写入校验失败：{path.name} 工作表 {name} 行数 {len(got)} != 预期 {len(frame)}"
                )
    os.replace(tmp, path)


def atomic_write_csv(path: Path, frame: pd.DataFrame) -> None:
    assert_writable(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f".{path.stem}.tmp{path.suffix}")
    frame.to_csv(tmp, index=False, encoding="utf-8-sig")
    os.replace(tmp, path)


def write_stream_sheet(book, sheet_name: str, frame: pd.DataFrame, head_fmt=None,
                       chunk: int = 100000):
    """Row-major streaming write for xlsxwriter constant_memory sheets.

    pandas `to_excel` emits cells column by column, which constant_memory silently
    discards, so large sheets must be written row by row here instead.
    """
    ws = book.add_worksheet(sheet_name[:31])
    cols = list(frame.columns)
    for idx, col in enumerate(cols):
        ws.write(0, idx, str(col), head_fmt)
        ws.set_column(idx, idx, min(max(len(str(col)) + 4, 12), 40))
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(frame), 1), max(len(cols) - 1, 0))
    if frame.empty:
        return ws

    col_arrays = []
    for col in cols:
        series = frame[col]
        if pd.api.types.is_datetime64_any_dtype(series):
            formatted = series.dt.strftime("%Y-%m-%d %H:%M:%S")
            arr = formatted.astype(object).where(series.notna(), None).to_numpy()
        else:
            arr = series.astype(object).where(series.notna(), None).to_numpy()
        col_arrays.append(arr)

    write_row = ws.write_row
    n = len(frame)
    for start in range(0, n, chunk):
        end = min(start + chunk, n)
        slices = [arr[start:end] for arr in col_arrays]
        for row_offset, values in enumerate(zip(*slices), start=start + 1):
            write_row(row_offset, 0, values)
    return ws


def write_workbook(path: Path, sheets: dict[str, pd.DataFrame], font: str | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter

    book = xlsxwriter.Workbook(str(path), {"constant_memory": True,
                                           "default_date_format": "yyyy-mm-dd hh:mm:ss"})
    fmt = {"bold": True, "bg_color": "#DDEBF7", "border": 1, "align": "center"}
    if font:
        fmt["font_name"] = font
    head = book.add_format(fmt)
    for name, frame in sheets.items():
        write_stream_sheet(book, name, frame, head)
    book.close()


def write_status_check_list(path: Path, tnos, font: str | None = None) -> None:
    """One-column Excel for external status query: header Tracking Number, then one TNO per row."""
    import xlsxwriter

    path.parent.mkdir(parents=True, exist_ok=True)
    values = [str(v).strip() for v in tnos if v is not None and str(v).strip() and str(v).strip().lower() != "nan"]
    book = xlsxwriter.Workbook(str(path))
    ws = book.add_worksheet("Tracking Number")
    head = {"bold": True, "font_name": font or "微软雅黑", "num_format": "@"}
    cell = {"font_name": font or "微软雅黑", "num_format": "@"}
    ws.write_string(0, 0, "Tracking Number", book.add_format(head))
    ws.set_column(0, 0, 32, book.add_format(cell))
    for row, tno in enumerate(values, start=1):
        ws.write_string(row, 0, tno)
    book.close()


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))
