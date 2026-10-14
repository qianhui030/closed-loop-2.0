# -*- coding: utf-8 -*-
"""No Ascan Pool persistence. The pool is long-lived tracking data, never rebuilt empty."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from common import (
    DIRS,
    POOL_ACTIVE_PARQUET,
    POOL_ACTIVE_SHEET,
    POOL_CLOSED_PARQUET,
    POOL_CLOSED_SHEET,
    POOL_FILE,
    StopWithMessage,
    assert_writable,
    atomic_write_excel,
    atomic_write_parquet,
    backup_file,
)

# Business fields carried over from the Unified Master wide table.
POOL_BUSINESS_COLS = [
    "TNO", "Source", "original_source", "File名称", "primary_source", "业务名称",
    "MAWB", "CUSTOMER", "customer_raw", "Hub", "start_time", "start_time_type",
    "first_ascan_time", "Ascan Due Date", "cohort_date", "sla_hours", "sla_type", "status",
    "跨Hub转运", "has_ags", "has_system", "has_platform", "platform_sheet",
    "ags_mawb_hub", "ags_destination_hub", "hub_source", "start_time_source",
]
POOL_EXTRA_COLS = [
    "pool_entry_date", "receiving_time", "receiving_record_id",
    "kpi_start_time", "kpi_start_time_source", "kpi_due_date", "kpi_date",
    "kpi_deadline",
    "pool_state", "kpi_counted", "counted_kpi_date", "counted_run_id",
    "last_update_time", "closed_time",
]
POOL_AUDIT_COLS = ["receiving_hub", "hub_match_status", "match_status", "pool_note"]
POOL_COLS = POOL_BUSINESS_COLS + POOL_EXTRA_COLS + POOL_AUDIT_COLS

DT_COLS = [
    "start_time", "first_ascan_time", "Ascan Due Date",
    "receiving_time", "kpi_start_time", "kpi_due_date", "kpi_deadline",
    "last_update_time", "closed_time",
]

# pool_state values
STATE_TRACKING = "Tracking"        # 在池中：未到期（可已有 A-scan）或仍在等 A-scan
STATE_PENDING = "Pending_KPI"      # 第一步已转入本次考核，等第二步确认
STATE_CLOSED = "Closed"            # 已完成考核与追踪


def empty_pool() -> pd.DataFrame:
    return pd.DataFrame(columns=POOL_COLS)


def _coerce(df: pd.DataFrame) -> pd.DataFrame:
    for col in POOL_COLS:
        if col not in df.columns:
            df[col] = pd.NA
    for col in DT_COLS:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    # An all-empty column read back from Excel comes in as float64 and then rejects
    # text assignment, so pin every non-datetime pool column to object.
    for col in POOL_COLS:
        if col not in DT_COLS and col != "sla_hours":
            df[col] = df[col].astype("object")
    # Excel round-trips sla_hours as float64; master/KPI compute often emits Int64.
    # Keep one numeric dtype so upsert does not raise on NA + 12/24/48.
    df["sla_hours"] = pd.to_numeric(df["sla_hours"], errors="coerce").astype("float64")
    df["TNO"] = df["TNO"].astype("string").str.strip()
    return df[POOL_COLS + [c for c in df.columns if c not in POOL_COLS]]


def pool_store_exists() -> bool:
    return POOL_ACTIVE_PARQUET.exists() or POOL_CLOSED_PARQUET.exists() or POOL_FILE.exists()


def _read_parquet(path: Path) -> pd.DataFrame:
    if not path.exists():
        return empty_pool()
    return pd.read_parquet(path)


def load_pool(logger) -> tuple[pd.DataFrame, pd.DataFrame, bool]:
    """Returns (active, closed, first_time). Never fabricates history from an old master."""
    if not pool_store_exists():
        logger.warning("=" * 70)
        logger.warning("Pool 不存在，已创建空模板。")
        logger.warning("历史 Pool 尚未导入：本次没有任何历史未 A-scan 记录参与考核。")
        logger.warning("如果确实有历史 Pool，请先把它放到 %s 再重跑。", DIRS["pool_dir"])
        logger.warning("程序不会把任意旧主表自动当作未考核历史数据。")
        logger.warning("=" * 70)
        active, closed = empty_pool(), empty_pool()
        save_pool(active, closed, logger, reason="create_empty_template")
        return active, closed, True

    if POOL_ACTIVE_PARQUET.exists() or POOL_CLOSED_PARQUET.exists():
        logger.info("读取 Pool parquet")
        active = _coerce(_read_parquet(POOL_ACTIVE_PARQUET).copy())
        closed = _coerce(_read_parquet(POOL_CLOSED_PARQUET).copy())
    else:
        assert_writable(POOL_FILE)
        logger.info("读取 Pool Excel（将迁移为 parquet）")
        book = pd.read_excel(POOL_FILE, sheet_name=None, engine="openpyxl",
                             dtype={"TNO": "string", "MAWB": "string"})
        active = _coerce(book.get(POOL_ACTIVE_SHEET, empty_pool()).copy())
        closed = _coerce(book.get(POOL_CLOSED_SHEET, empty_pool()).copy())
    if active["TNO"].duplicated().any():
        dup = active.loc[active["TNO"].duplicated(keep=False), "TNO"].unique()[:10]
        raise StopWithMessage(
            f"Active_Pool 存在重复 TNO（每个 TNO 只能有一条当前记录）。示例：{list(dup)}"
        )
    logger.info("读取 Pool：Active %d 条，Closed_History %d 条", len(active), len(closed))
    return active, closed, False


def _archive_closed(closed: pd.DataFrame, cfg: dict, logger) -> pd.DataFrame:
    """Split Closed_History by date once it approaches the Excel row limit. No truncation."""
    limit = int(cfg.get("pool", {}).get("closed_history_max_rows", 900000))
    if len(closed) <= limit:
        return closed
    keys = pd.to_datetime(closed["counted_kpi_date"], errors="coerce")
    keys = keys.dt.strftime("%Y-%m").fillna("unknown")
    order = sorted(keys.unique())
    keep_mask = pd.Series(True, index=closed.index)
    for month in order:
        if int(keep_mask.sum()) <= limit:
            break
        part = closed[keys == month]
        out = DIRS["pool_dir"] / f"Closed_History_{month}.xlsx"
        atomic_write_excel(out, {POOL_CLOSED_SHEET: part})
        logger.warning("Closed_History 超过 %d 行，已按日期归档 %s（%d 条），未截断。",
                       limit, out.name, len(part))
        keep_mask &= keys != month
    return closed[keep_mask].copy()


def save_pool(active: pd.DataFrame, closed: pd.DataFrame, logger,
              cfg: dict | None = None, reason: str = "") -> None:
    cfg = cfg or {}
    pool_cfg = cfg.get("pool") or {}
    keep = int(pool_cfg.get("backup_keep", 30))
    closed = _archive_closed(_coerce(closed.copy()), cfg, logger)
    active = _coerce(active.copy())
    for path in (POOL_ACTIVE_PARQUET, POOL_CLOSED_PARQUET):
        backup = backup_file(path, DIRS["pool_backups"], keep=keep)
        if backup:
            logger.info("Pool 更新前备份：%s", backup.name)
    atomic_write_parquet(POOL_ACTIVE_PARQUET, active[POOL_COLS])
    atomic_write_parquet(POOL_CLOSED_PARQUET, closed[POOL_COLS])
    write_excel = bool(pool_cfg.get("write_excel", False))
    if write_excel:
        atomic_write_excel(POOL_FILE, {
            POOL_ACTIVE_SHEET: active[POOL_COLS],
            POOL_CLOSED_SHEET: closed[POOL_COLS],
        }, verify=False)
        logger.info("Pool Excel 副本已写入 %s", POOL_FILE.name)
    logger.info("Pool 已保存%s：Active %d 条，Closed_History %d 条（parquet）",
                f"（{reason}）" if reason else "", len(active), len(closed))


def upsert_active(active: pd.DataFrame, rows: pd.DataFrame, logger) -> pd.DataFrame:
    """Update by TNO, never append a second row for the same TNO."""
    if rows.empty:
        return active
    rows = _coerce(rows.copy())
    if active.empty:
        return rows.drop_duplicates("TNO", keep="last").reset_index(drop=True)
    active = active.set_index("TNO", drop=False)
    incoming = rows.drop_duplicates("TNO", keep="last").set_index("TNO", drop=False)
    common = incoming.index.intersection(active.index)
    fresh = incoming.index.difference(active.index)
    for col in POOL_COLS:
        if col in incoming.columns and len(common):
            new_vals = incoming.loc[common, col]
            old_vals = active.loc[common, col]
            # Empty incoming values never overwrite an existing confirmed value.
            filler = new_vals.where(new_vals.notna(), old_vals)
            if col == "sla_hours":
                active.loc[common, col] = pd.to_numeric(filler, errors="coerce").astype("float64")
            else:
                active.loc[common, col] = filler.values
    out = pd.concat([active, incoming.loc[fresh]], ignore_index=True)
    logger.info("Active_Pool upsert：更新 %d 条，新增 %d 条", len(common), len(fresh))
    return _coerce(out)
