# -*- coding: utf-8 -*-
"""Bridge to Data Combination. It stays the only implementation of the 3-source merge."""
from __future__ import annotations

import importlib.util
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from common import DIRS, ROOT, StopWithMessage, clean_text

_MODULE = None

TEXT_COLS = ["TNO", "MAWB", "ags_destination_hubs"]
DT_COLS = [
    "start_time", "first_ascan_time", "Ascan Due Date",
    "交接时间", "Ascan时间", "ags_handover_time", "ags_ascan_time",
]


def merge_module():
    """Import build_unified_master so SLA/customer rules have exactly one definition."""
    global _MODULE
    if _MODULE is not None:
        return _MODULE
    path = DIRS["data_combination"] / "build_unified_master.py"
    if not path.exists():
        raise StopWithMessage(f"找不到三源合并程序：{path}")
    spec = importlib.util.spec_from_file_location("build_unified_master", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules["build_unified_master"] = module
    spec.loader.exec_module(module)
    _MODULE = module
    return module


def wide_csv_path(cfg: dict) -> Path:
    section = cfg["unified_master"]
    return DIRS["data_combination"] / section["wide_csv"]


def wide_parquet_path(cfg: dict) -> Path:
    csv_path = wide_csv_path(cfg)
    return csv_path.with_suffix(".parquet")


def run_merge(cfg: dict, logger) -> str:
    """Run build_unified_master.py in its own folder and surface its sheet report."""
    folder = DIRS["data_combination"]
    entry = cfg["unified_master"]["merge_entry"]
    logger.info("调用三源合并：%s / %s", folder.name, entry)
    proc = subprocess.run(
        [sys.executable, entry],
        cwd=folder,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    out = (proc.stdout or "") + (proc.stderr or "")
    for line in out.splitlines():
        logger.info("  [merge] %s", line)
    if proc.returncode != 0:
        raise StopWithMessage(
            f"三源合并失败（returncode={proc.returncode}）。已中止，未生成任何正式结果。\n"
            f"最后输出：\n{out[-2000:]}"
        )
    return out


def check_platform_sheet_report(merge_log: str, logger) -> list[str]:
    """Pull the platform tab disposition out of the merge log so nothing drops silently."""
    notes = []
    capture = False
    for line in merge_log.splitlines():
        stripped = line.strip()
        if stripped.startswith("-- 已识别明细表 --"):
            capture = True
        if capture:
            notes.append(stripped)
        if stripped.startswith("platform combined unique TNO"):
            notes.append(stripped)
            capture = False
    absent = [n for n in notes if "业务缺席" in n]
    if absent:
        logger.warning("平台业务缺席（本批次源文件里确实没有这些明细表）：")
        for item in absent:
            logger.warning("  %s", item)
    if "Missing source files this run" in merge_log:
        logger.warning("本批次三源有缺席文件，已用现有文件继续合并。缺席来源的包裹不会进入当天三源主表。")
    return notes


def load_unified_master(cfg: dict, mode: str, allow_stale: bool, logger) -> tuple[pd.DataFrame, dict]:
    """mode='run' 重跑合并；mode='reuse' 使用本次已生成的宽表。不静默读旧结果。"""
    csv_path = wide_csv_path(cfg)
    pq_path = wide_parquet_path(cfg)
    merge_log = ""
    if mode == "run":
        merge_log = run_merge(cfg, logger)
        check_platform_sheet_report(merge_log, logger)

    # Prefer the freshest of parquet and csv so both fast runs and legacy/test CSVs work.
    candidates = [p for p in (pq_path, csv_path) if p.exists()]
    path = max(candidates, key=lambda p: p.stat().st_mtime) if candidates else pq_path

    if mode == "reuse":
        if not path.exists():
            raise StopWithMessage(
                f"找不到宽表 {path}。\n请改用 --merge run 先跑一次三源合并。"
            )
        age = datetime.now() - datetime.fromtimestamp(path.stat().st_mtime)
        limit = int(cfg["unified_master"].get("reuse_max_age_minutes", 180))
        logger.info("复用已有宽表：%s（生成于 %s，%d 分钟前）",
                    path.name,
                    datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    int(age.total_seconds() // 60))
        if age > timedelta(minutes=limit) and not allow_stale:
            raise StopWithMessage(
                f"宽表 {path.name} 生成于 {int(age.total_seconds() // 60)} 分钟前，"
                f"超过 {limit} 分钟阈值。\n"
                "这可能是上次遗留的输出。请选择其一：\n"
                "  1) --merge run              重新跑三源合并\n"
                "  2) --merge reuse --allow-stale-master   明确确认就要用这份旧宽表"
            )
    elif mode != "run":
        raise StopWithMessage(f"未知的 --merge 取值：{mode}（只能是 run 或 reuse）")

    if not path.exists():
        raise StopWithMessage(f"三源合并结束但没有生成宽表：{path}")

    logger.info("读取宽表 %s", path)
    if path.suffix.lower() == ".parquet":
        df = pd.read_parquet(path)
        for col in TEXT_COLS:
            if col in df.columns:
                df[col] = df[col].astype("string")
    else:
        df = pd.read_csv(path, dtype={c: "string" for c in TEXT_COLS}, low_memory=False)
    for col in DT_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    tno = df["TNO"].astype("string").str.strip()
    tno = tno.where(~tno.str.lower().isin({"", "nan", "none", "<na>", "nat", "null"}), pd.NA)
    df["TNO"] = tno.astype("object").where(tno.notna(), None)
    before = len(df)
    df = df[df["TNO"].notna()].copy()
    if len(df) != before:
        logger.warning("宽表中有 %d 行 TNO 为空，已排除并记录", before - len(df))
    if df["TNO"].duplicated().any():
        dup = int(df["TNO"].duplicated().sum())
        raise StopWithMessage(f"宽表存在重复 TNO {dup} 行，闭环拒绝继续。请先修复三源合并输出。")

    meta = {
        "wide_csv": str(path.relative_to(ROOT)),
        "wide_csv_mtime": datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        "wide_csv_rows": int(len(df)),
        "merge_mode": mode,
        "source_files": sorted({s for s in df["source_files"].dropna().astype(str)
                                for s in s.split(";")})[:20],
    }
    logger.info("宽表 %d 行，唯一 TNO %d", len(df), df["TNO"].nunique())
    return df, meta
