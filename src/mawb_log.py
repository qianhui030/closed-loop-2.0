# -*- coding: utf-8 -*-
"""Daily unique-MAWB registry. One row per MAWB, appended across KPI dates."""
from __future__ import annotations

import pandas as pd

from common import (
    DIRS,
    clean_text,
    join_unique,
    atomic_write_excel,
)

MAWB_SHEET = "MAWB"
MAWB_LOG_COLS = ["Source", "kpi_date", "Hub", "MAWB", "start_time"]
MAWB_SUMMARY_COLS = [
    "Hub", "CUSTOMER", "MAWB", "Source", "跨Hub转运", "start_time",
    "Count of TNO", "Sum of onTime_count", "Average of onTime_count",
    "Sum of completed_count", "Completion percentage",
]
CROSS_HUB_BLANK = {"", "nan", "none", "<na>", "nat", "null", "n/a", "非跨Hub", "非跨hub"}


def history_path():
    return DIRS["mawb_log"] / "MAWB_History.xlsx"


def _norm_mawb(series: pd.Series) -> pd.Series:
    return series.map(clean_text).astype("string")


def _hub_label(value) -> str:
    text = str(value).strip() if value is not None and pd.notna(value) else ""
    if not text or text in {"(空Hub)", "<NA>", "nan", "None"}:
        return "(空Hub)"
    return text


def _pick_customer(grp: pd.DataFrame) -> str:
    """Updated Unified Master CUSTOMER only. Blank stays empty; never customer_raw."""
    if "CUSTOMER" not in grp.columns:
        return "(空客户)"
    return join_unique(grp["CUSTOMER"]) or "(空客户)"


def _pick_cross_hub(grp: pd.DataFrame) -> str | None:
    """Keep LAX-DFW / EWR-ORD (and Found). Blank or 非跨Hub stay empty."""
    if "跨Hub转运" not in grp.columns:
        return None
    kept: list[str] = []
    for raw in grp["跨Hub转运"]:
        text = clean_text(raw)
        if text is None or text.lower() in CROSS_HUB_BLANK:
            continue
        if text not in kept:
            kept.append(text)
    return ";".join(kept) if kept else None


def _completed_mask(grp: pd.DataFrame) -> pd.Series:
    """Completed = On Time or any Late*. Unscanned / 51005 / Missing Ascan do not count."""
    if "status" in grp.columns:
        st = grp["status"].astype("string")
    elif "kpi_status" in grp.columns:
        st = grp["kpi_status"].astype("string")
    else:
        ascan = pd.to_datetime(grp.get("first_ascan_time"), errors="coerce")
        return ascan.notna()
    return st.eq("On Time") | st.str.startswith("Late", na=False)


def summarize_mawbs(eligible: pd.DataFrame, kpi_date: str, hub_order: list[str]) -> pd.DataFrame:
    """One row per Hub + MAWB. A master bill in two hubs keeps both rows."""
    empty = pd.DataFrame(columns=MAWB_LOG_COLS + [
        "CUSTOMER", "Count of TNO", "Sum of onTime_count", "Average of onTime_count",
        "Sum of completed_count", "Completion percentage",
    ])
    if eligible is None or eligible.empty or "MAWB" not in eligible.columns:
        return empty
    src = eligible.copy()
    src["MAWB"] = _norm_mawb(src["MAWB"])
    src = src[src["MAWB"].notna()].copy()
    if src.empty:
        return empty
    src["_hub"] = src["Hub"].map(_hub_label) if "Hub" in src.columns else "(空Hub)"

    rows = []
    for (mawb, hub), grp in src.groupby(["MAWB", "_hub"], sort=False):
        start = pd.to_datetime(grp["start_time"], errors="coerce").min() if "start_time" in grp.columns else pd.NaT
        one = grp.drop_duplicates("TNO") if "TNO" in grp.columns else grp
        ontime = pd.to_numeric(one.get("onTime_count", 0), errors="coerce").fillna(0)
        completed = int(_completed_mask(one).sum())
        n = int(one["TNO"].nunique()) if "TNO" in one.columns else len(one)
        source_col = "Source" if "Source" in one.columns else "original_source"
        rows.append({
            "Source": join_unique(one[source_col]) if source_col in one.columns else None,
            "kpi_date": kpi_date,
            "Hub": hub,
            "CUSTOMER": _pick_customer(one),
            "MAWB": str(mawb),
            "跨Hub转运": _pick_cross_hub(one),
            "start_time": start if pd.notna(start) else pd.NaT,
            "Count of TNO": n,
            "Sum of onTime_count": int(ontime.sum()),
            "Average of onTime_count": round(float(ontime.mean()), 6) if len(ontime) else 0.0,
            "Sum of completed_count": completed,
            "Completion percentage": round(completed / n, 6) if n else 0.0,
        })
    out = pd.DataFrame(rows)
    rank = {hub: i for i, hub in enumerate(hub_order)}
    out["_h"] = out["Hub"].map(lambda x: rank.get(x, 10_000))
    out = out.sort_values(["_h", "MAWB"], kind="mergesort").drop(columns="_h").reset_index(drop=True)
    return out


def summary_sheet(summary: pd.DataFrame) -> pd.DataFrame:
    if summary.empty:
        return pd.DataFrame(columns=MAWB_SUMMARY_COLS)
    return summary[MAWB_SUMMARY_COLS].copy()


def load_history() -> pd.DataFrame:
    path = history_path()
    if not path.exists():
        return pd.DataFrame(columns=MAWB_LOG_COLS)
    book = pd.read_excel(path, sheet_name=None, dtype={"MAWB": str, "kpi_date": str})
    frame = book.get(MAWB_SHEET)
    if frame is None:
        frame = next(iter(book.values()), pd.DataFrame())
    for col in MAWB_LOG_COLS:
        if col not in frame.columns:
            frame[col] = pd.NA
    frame["MAWB"] = _norm_mawb(frame["MAWB"])
    frame["kpi_date"] = frame["kpi_date"].astype("string")
    frame["Hub"] = frame["Hub"].astype("string")
    frame["Source"] = frame["Source"].astype("string")
    frame["start_time"] = pd.to_datetime(frame["start_time"], errors="coerce")
    return frame[MAWB_LOG_COLS]


def append_daily(summary: pd.DataFrame, logger, keep_backups: int = 0) -> pd.DataFrame:
    """Append this day's unique MAWBs into one MAWB_History.xlsx. Never duplicate, never copy the whole file."""
    incoming = summary[MAWB_LOG_COLS].copy() if len(summary) else pd.DataFrame(columns=MAWB_LOG_COLS)
    incoming["MAWB"] = _norm_mawb(incoming["MAWB"])
    incoming = incoming[incoming["MAWB"].notna()].copy()
    history = load_history()
    known = set(history["MAWB"].dropna().astype(str)) if len(history) else set()
    fresh = incoming[~incoming["MAWB"].astype(str).isin(known)].copy()
    merged = pd.concat([history, fresh], ignore_index=True) if len(fresh) else history
    if len(merged):
        merged["MAWB"] = merged["MAWB"].astype(str)
        dup = merged["MAWB"].duplicated()
        if dup.any():
            merged = merged.loc[~dup].copy()
    path = history_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    atomic_write_excel(path, {MAWB_SHEET: merged[MAWB_LOG_COLS]})
    logger.info("MAWB 累计表 %s（只此一份，追加写入）：历史 %d，本次新增 %d，跳过已有 %d",
                path.name, len(merged), len(fresh), int(len(incoming) - len(fresh)))
    return merged
