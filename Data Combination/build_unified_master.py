# -*- coding: utf-8 -*-
"""Build one-row-per-TNO unified master from AGS / System / Platform files.
Does not modify original Excel files.
"""
from __future__ import annotations

import math
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "output"

TIKTOK_CUSTOMERS = {"TIKTOK", "TIKTOK INC", "TT GS", "TT FBT"}
TIKTOK_RE = re.compile(r"tiktok", re.I)

# Display Customer aliases. Original legal names remain in conflict_audit.
CUSTOMER_ALIASES = {
    "HANGZHOU CAINIAO SUPPLY CHAIN MANAGEMENT CO.,LTD.": "AliExpress",
}

DISPLAY_COLS = [
    "Source", "File名称", "primary_source", "业务名称", "TNO", "MAWB",
    "CUSTOMER", "customer_raw", "Hub", "start_time", "start_time_type",
    "first_ascan_time", "Ascan Due Date", "cohort_date", "time_diff_hours",
    "sla_hours", "sla_type", "status", "跨Hub转运",
]

SLA_TYPE_HOURS = "hours"
SLA_TYPE_CALENDAR_D0 = "calendar_d0"
SLA_TYPE_CALENDAR_D1 = "calendar_d1"
CALENDAR_SLA_TYPES = (SLA_TYPE_CALENDAR_D0, SLA_TYPE_CALENDAR_D1)
DISPLAY_DUE_TIME = pd.Timedelta(hours=23, minutes=59, seconds=59)
DATETIME_COLS = ["start_time", "first_ascan_time", "Ascan Due Date"]
AGS_EXTRA_COLS = [
    "ags_mawb_hub", "ags_tracking_origin_hub", "ags_destination_hub",
    "ags_destination_facility", "ags_handover_time", "ags_ascan_time",
    "ags_destination_hubs", "跨Hub转运",
]
INTER_HUB_RECALC_COLS = [
    "actual_receiving_time", "actual_receiving_time_source",
    "recalculated_start_time", "recalculated_ascan_due_date",
    "recalculated_time_diff_hours", "recalculated_status",
]
LAX_TEXAS_DESTS = {"DFW", "IAH"}
EWR_DOWNSTREAM_DESTS = {"ATL", "EWR", "MIA", "ORD"}
FULL_COLS = DISPLAY_COLS + [
    "Customer", "customer_standardized", "交接时间", "交接时间类型", "Ascan时间",
    "start_time_type_raw", "within_24h", "monitor_flag", "merge_note",
    "source_files", "source_count", "has_ags", "has_system", "has_platform",
    "customer_source", "hub_source", "start_time_source", "ascan_source", "mawb_source",
    "mawb_mapping_conflict_flag", "platform_sheet",
] + [c for c in AGS_EXTRA_COLS if c != "跨Hub转运"]


def is_blank_series(s: pd.Series) -> pd.Series:
    if s.dtype == object or str(s.dtype) == "string":
        t = s.astype("string").str.strip()
        return s.isna() | t.isin(["", "nan", "None", "NaT", "<NA>", "NaN", "none", "NULL"])
    return s.isna()


def clean_tn_series(s: pd.Series) -> pd.Series:
    out = []
    for v in s:
        if v is None:
            out.append(None)
            continue
        try:
            if pd.isna(v):
                out.append(None)
                continue
        except Exception:
            pass
        if isinstance(v, (int, np.integer)):
            out.append(str(int(v)))
            continue
        if isinstance(v, (float, np.floating)):
            if math.isnan(v) or math.isinf(v):
                out.append(None)
                continue
            if float(v).is_integer():
                out.append(str(int(v)))
                continue
        t = str(v).strip()
        if t.lower() in {"", "nan", "none", "nat", "null", "<na>"}:
            out.append(None)
            continue
        if re.fullmatch(r"\d+\.0+", t):
            t = t.split(".")[0]
        out.append(t)
    return pd.Series(out, index=s.index, dtype="object")


def clean_mawb_series(s: pd.Series) -> pd.Series:
    t = s.astype("string").str.strip()
    t = t.replace({"<NA>": pd.NA, "nan": pd.NA, "None": pd.NA, "": pd.NA})
    return t


def parse_dt_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(pd.NaT, index=s.index if s is not None else None)
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")
    x = s.astype("string").str.strip()
    x = x.str.replace(r"\.0$", "", regex=True)
    x = x.replace({"<NA>": pd.NA, "nan": pd.NA, "None": pd.NA, "NaT": pd.NA, "": pd.NA})
    return pd.to_datetime(x, errors="coerce")


def fmt_dt(s: pd.Series) -> pd.Series:
    dt = parse_dt_series(s) if not pd.api.types.is_datetime64_any_dtype(s) else s
    out = dt.dt.strftime("%Y-%m-%d %H:%M:%S")
    return out.astype("object").where(dt.notna(), None)


FBT_12H_HUBS = {"MIA"}


def is_fbt_row(df: pd.DataFrame) -> pd.Series:
    """TT FBT identity. SLA is 12h only at Hub MIA; other hubs are 24h.
    Do not treat Winit Manifest Time display as FBT."""
    if "CUSTOMER" in df.columns:
        cust = df["CUSTOMER"].astype("string").str.strip()
    elif "Customer" in df.columns:
        cust = df["Customer"].astype("string").str.strip()
    else:
        cust = pd.Series("", index=df.index, dtype="string")
    biz = df["业务名称"].astype("string") if "业务名称" in df.columns else pd.Series("", index=df.index)
    plat = df["platform_sheet"].astype("string") if "platform_sheet" in df.columns else pd.Series("", index=df.index)
    return (
        cust.str.casefold().eq("tt fbt").fillna(False)
        | biz.str.contains("FBT", na=False)
        | plat.eq("FBT").fillna(False)
    )


def is_fbt_12h_row(df: pd.DataFrame) -> pd.Series:
    """12h only for TT FBT whose final Hub is MIA. Other TT FBT hubs stay 24h."""
    fbt = is_fbt_row(df)
    if "Hub" not in df.columns:
        return pd.Series(False, index=df.index)
    hub = df["Hub"].astype("string").str.strip().str.upper()
    return fbt & hub.isin(FBT_12H_HUBS)


def remap_fbt_dfw_hub(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """DFW has no TT FBT. Those parcels are IAH."""
    if df is None or df.empty or "Hub" not in df.columns:
        return df, 0
    hub = df["Hub"].astype("string").str.strip().str.upper()
    hit = is_fbt_row(df) & hub.eq("DFW")
    n = int(hit.sum())
    if n == 0:
        return df, 0
    out = df.copy()
    out.loc[hit, "Hub"] = "IAH"
    if "merge_note" in out.columns:
        note = out["merge_note"].astype("string").fillna("")
        out.loc[hit, "merge_note"] = (note.loc[hit] + ";fbt_dfw_to_iah").str.strip(";")
    return out, n


def _customer_blob(df: pd.DataFrame) -> pd.Series:
    parts = []
    for col in ("CUSTOMER", "Customer", "业务名称", "customer_raw"):
        if col in df.columns:
            t = df[col].astype("string").str.lower().str.strip()
            t = t.str.replace(r"[\s_\-]+", " ", regex=True)
            parts.append(t.fillna(""))
    if not parts:
        return pd.Series("", index=df.index, dtype="string")
    blob = parts[0]
    for p in parts[1:]:
        blob = blob + ";" + p
    return blob


def is_48h_sla_row(df: pd.DataFrame) -> pd.Series:
    """+48h only for System-only ProCarrier / Vinted US Domestic. AGS stays +24h."""
    blob = _customer_blob(df)
    name_hit = (
        blob.str.contains(r"pro\s*carrier", na=False)
        | blob.str.contains(r"vinted\s+us\s+domestic", na=False)
    )
    if "has_system" in df.columns:
        is_sys = df["has_system"].astype("string") == "Yes"
    elif "Source" in df.columns:
        is_sys = df["Source"].astype("string").str.contains("System", na=False)
    else:
        is_sys = pd.Series(False, index=df.index)
    if "has_ags" in df.columns:
        is_ags = df["has_ags"].astype("string") == "Yes"
    elif "Source" in df.columns:
        is_ags = df["Source"].astype("string").str.contains("AGS", na=False)
    else:
        is_ags = pd.Series(False, index=df.index)
    return name_hit & is_sys & ~is_ags


def display_start_type(stype: pd.Series) -> pd.Series:
    """Winit source field remains date_manifested; master display is Manifest Time."""
    t = stype.astype("string")
    return t.where(t != "date_manifested", "Manifest Time")


def _first_existing(df: pd.DataFrame, names: list[str]) -> pd.Series:
    for name in names:
        if name in df.columns:
            return df[name]
    return pd.Series(pd.NA, index=df.index)


def is_winit_row(df: pd.DataFrame) -> pd.Series:
    """Winit platform / customer rows. Exact name match; do not treat other businesses as Winit."""
    hit = pd.Series(False, index=df.index)
    for col in ("CUSTOMER", "Customer", "业务名称", "platform_sheet"):
        if col in df.columns:
            hit = hit | df[col].astype("string").str.strip().str.casefold().eq("winit")
    return hit.fillna(False)


def is_winit_lax_pickup_row(df: pd.DataFrame) -> pd.Series:
    """LAX 自提 is D0. Hub must already say 自提; ordinary Winit LAX stays D1."""
    if "Hub" not in df.columns:
        return pd.Series(False, index=df.index)
    hub = df["Hub"].astype("string").str.strip()
    return is_winit_row(df) & hub.str.contains("自提", na=False)


def compute_sla_fields(df: pd.DataFrame, start: pd.Series) -> pd.DataFrame:
    """One SLA definition for merge and closed-loop.

    Hour customers: due = start + 12/24/48h (inclusive).
    TT FBT Hub MIA is 12h; every other TT FBT hub is 24h.
    Winit D1: due date = manifest local date + 1 calendar day; no weekend skip.
    Winit LAX自提 D0: due date = manifest local date.
    Calendar display due is 23:59:59; program deadline is next day 00:00:00 exclusive.
    """
    start = pd.to_datetime(start, errors="coerce")
    fbt = is_fbt_row(df)
    fbt12 = is_fbt_12h_row(df)
    winit = is_winit_row(df)
    pickup = is_winit_lax_pickup_row(df)
    sla48 = is_48h_sla_row(df) & ~fbt & ~winit

    sla_type = pd.Series(SLA_TYPE_HOURS, index=df.index, dtype="object")
    sla_type = sla_type.mask(winit & pickup, SLA_TYPE_CALENDAR_D0)
    sla_type = sla_type.mask(winit & ~pickup, SLA_TYPE_CALENDAR_D1)

    hours = pd.Series(pd.array([24] * len(df), dtype="Int64"), index=df.index)
    hours = hours.mask(fbt12, 12)
    hours = hours.mask(sla48, 48)
    hours = hours.mask(winit, pd.NA)
    if "sla_hours" in df.columns:
        existing = pd.to_numeric(df["sla_hours"], errors="coerce")
        # Recompute FBT. A stored 12h must not keep non-MIA TT FBT on the old SLA.
        keep = existing.notna() & ~winit & ~fbt
        hours = hours.where(~keep, existing.round().astype("Int64"))

    cal = sla_type.isin(CALENDAR_SLA_TYPES)
    start_day = start.dt.normalize()
    cal_due_day = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    cal_due_day = cal_due_day.mask(sla_type.eq(SLA_TYPE_CALENDAR_D1), start_day + pd.Timedelta(days=1))
    cal_due_day = cal_due_day.mask(sla_type.eq(SLA_TYPE_CALENDAR_D0), start_day)

    hour_due = start + pd.to_timedelta(hours.astype("Float64"), unit="h")
    due_display = hour_due.copy()
    due_display = due_display.mask(cal, cal_due_day + DISPLAY_DUE_TIME)

    deadline = hour_due.copy()
    deadline = deadline.mask(cal, cal_due_day + pd.Timedelta(days=1))

    kpi_date = pd.Series(pd.NA, index=df.index, dtype="object")
    kpi_date = kpi_date.mask(cal & cal_due_day.notna(), cal_due_day.dt.strftime("%Y-%m-%d"))
    kpi_date = kpi_date.mask(~cal & hour_due.notna(), hour_due.dt.strftime("%Y-%m-%d"))

    return pd.DataFrame({
        "sla_type": sla_type,
        "sla_hours": hours,
        "due_display": due_display,
        "deadline": deadline,
        "kpi_date": kpi_date,
        "calendar": cal,
    }, index=df.index)


def ascan_on_time(ascan: pd.Series, deadline: pd.Series, calendar: pd.Series) -> pd.Series:
    """Hour SLA: ascan <= deadline. Calendar SLA: ascan < next-day 00:00:00."""
    ascan = pd.to_datetime(ascan, errors="coerce")
    deadline = pd.to_datetime(deadline, errors="coerce")
    cal = calendar.fillna(False).astype(bool)
    comparable = ascan.notna() & deadline.notna()
    return comparable & ((~cal & (ascan <= deadline)) | (cal & (ascan < deadline)))


def add_ascan_due(df: pd.DataFrame) -> pd.DataFrame:
    """Due from local start_time. Winit is calendar D0/D1; others remain hour-based."""
    start = parse_dt_series(_first_existing(df, ["start_time", "交接时间"]))
    ascan = parse_dt_series(_first_existing(df, ["first_ascan_time", "Ascan时间"]))
    spec = compute_sla_fields(df, start)
    due = spec["due_display"]
    df = df.copy()
    df["sla_hours"] = spec["sla_hours"]
    df["sla_type"] = spec["sla_type"]
    df["Ascan Due Date"] = due
    df["start_time"] = start
    df["first_ascan_time"] = ascan
    df["交接时间"] = start
    df["Ascan时间"] = ascan
    status = pd.Series(pd.NA, index=df.index, dtype="object")
    has_ascan = ascan.notna()
    on_time = ascan_on_time(ascan, spec["deadline"], spec["calendar"])
    late = has_ascan & spec["deadline"].notna() & ~on_time
    status = status.mask(on_time, "On Time")
    status = status.mask(late, "Late")
    status = status.mask(~has_ascan, "No A-Scan")
    df["status"] = status
    return df


def mask_tn(tn: str) -> str:
    if not tn:
        return ""
    tn = str(tn)
    if len(tn) <= 8:
        return tn[:2] + "***" + tn[-2:]
    return tn[:6] + "****" + tn[-4:]


def _norm_sheet(name: str) -> str:
    return re.sub(r"[\s_\-]+", "", str(name).strip()).lower()


def _norm_col(name: str) -> str:
    return re.sub(r"\s+", "", str(name).strip()).lower()


# Exact names plus dated/renamed tabs. Pivot/summary tabs are excluded before matching.
PLATFORM_SHEET_ALIASES = {
    "TT-GS": [
        "TT-GS明细 ", "TT-GS明细", "TT GS明细", "TTGS明细", "TTGS", "GS明细",
        "TT-GS ASCAN数据", "TT-GS ASCAN 数据",
        "Parcels Detail", "TTGSParcels Detail",
    ],
    "FBT": [
        "TT-FBT明细", "FBT明细", "TT FBT明细", "TTFBT明细",
        "TT-FBT ASCAN 数据", "TT-FBT ASCAN数据", "TT-FBT ASCAN 数据",
    ],
    "Winit": [
        "Winit明细", "winit明细",
        "winit ASACN数据", "winit ASCAN数据", "winit ASCAN 数据",
    ],
    "AliExpress": [
        "Aliexpress-明细", "AliExperess明细", "AliExpress明细", "Aliexpress明细",
        "AliExpress ASACN数据", "AliExpress ASCAN数据", "AliExpress ASCAN 数据",
        "菜鸟明细",
        "cainiao 明细",
        "明细",
        "Ali明细",
    ],
}
PLATFORM_SHEET_KEYWORDS = {
    "TT-GS": {"must": ["ttgs"], "reject": ["ttgspv"]},
    "FBT": {"must": ["ttfbt", "fbt明细"], "reject": ["ttgs"]},
    "Winit": {"must": ["winit"], "reject": []},
    "AliExpress": {"must": ["aliexpress", "aliexperess", "cainiao", "ali明细"], "reject": []},
}
TRACKING_HEADER_NAMES = {
    _norm_col(x) for x in ("追踪号", "跟踪号", "运单号", "Tracking Number", "tracking_number", "Tracking No")
}


def _is_pivot_sheet(name: str) -> bool:
    """Pivot/summary/未上网 tabs. 'pv' as a standalone token, or glued on the end (TTGSPV)."""
    text = str(name).strip()
    if any(tok in text for tok in ("数透", "汇总", "透视", "未上网")):
        return True
    if re.search(r"(?<![A-Za-z])pv(?![A-Za-z])", text, flags=re.I):
        return True
    return _norm_sheet(text).endswith("pv")


def read_excel_fast(path: Path, **kwargs) -> pd.DataFrame:
    """Prefer python-calamine (Rust) for 10x+ faster Excel reading; fall back to openpyxl."""
    try:
        return pd.read_excel(path, engine="calamine", **kwargs)
    except Exception:
        return pd.read_excel(path, engine="openpyxl", **kwargs)


def read_sheet_headers(path: Path) -> dict[str, list[str]]:
    try:
        from python_calamine import CalamineWorkbook
        wb = CalamineWorkbook.from_path(str(path))
        headers = {}
        for name in wb.sheet_names:
            rows = wb.get_sheet_by_name(name).to_python(skip_empty_area=False)
            row = rows[0] if rows else ()
            headers[name] = [str(c).strip() for c in row if c is not None and str(c).strip()]
        return headers
    except Exception:
        wb = load_workbook(path, read_only=True, data_only=True)
        headers = {}
        for name in wb.sheetnames:
            row = next(wb[name].iter_rows(max_row=1, values_only=True), ())
            headers[name] = [str(c).strip() for c in row if c is not None and str(c).strip()]
        wb.close()
        return headers


def looks_like_parcel_detail(headers: list[str]) -> bool:
    return any(_norm_col(h) in TRACKING_HEADER_NAMES for h in headers)


def report_platform_sheets(path: Path, sheetnames: list[str], assigned: dict, headers: dict) -> None:
    """Print every tab's disposition, then fail on an unrecognized parcel detail tab."""
    excluded = [s for s in sheetnames if s not in assigned and _is_pivot_sheet(s)]
    others = [s for s in sheetnames if s not in assigned and s not in excluded]
    print("    -- 已识别明细表 --", flush=True)
    for sheet, key in assigned.items():
        print(f"      {key}: {sheet!r}", flush=True)
    if not assigned:
        print("      (无)", flush=True)
    print("    -- 已排除的汇总/数透/未上网表 --", flush=True)
    for sheet in excluded:
        print(f"      {sheet!r}", flush=True)
    if not excluded:
        print("      (无)", flush=True)
    print("    -- 未识别工作表 --", flush=True)
    for sheet in others:
        print(f"      {sheet!r} headers={headers.get(sheet, [])[:12]}", flush=True)
    if not others:
        print("      (无)", flush=True)
    suspects = [s for s in others if looks_like_parcel_detail(headers.get(s, []))]
    if suspects:
        detail = "; ".join(f"{s!r} headers={headers.get(s, [])[:12]}" for s in suspects)
        raise SystemExit(
            f"Unrecognized parcel detail sheet(s) in {path.name}: {detail}. "
            "Add an exact alias to PLATFORM_SHEET_ALIASES. "
            "Refusing to build a master that silently drops a business."
        )


def _skip_source_candidate(path: Path) -> bool:
    name = path.name
    lower = name.lower()
    if name.startswith("~$") or name.startswith("unified"):
        return True
    if "package_status" in lower or "status_check" in lower or "status check" in lower:
        return True
    return False


def _is_platform_workbook(path: Path, sheets: list[str]) -> bool:
    norms = {_norm_sheet(s) for s in sheets}
    for aliases in PLATFORM_SHEET_ALIASES.values():
        if any(_norm_sheet(alias) in norms for alias in aliases):
            return True
    name = path.name
    lower = name.lower()
    if "platform" in lower or "平台客户" in name or "ascan单号" in lower:
        return True
    return False


def platform_sheet_candidates(sheetnames: list[str], sheet_key: str) -> list[str]:
    """Detail tabs that could belong to this business. Pivot/summary tabs never qualify."""
    aliases = {_norm_sheet(a) for a in PLATFORM_SHEET_ALIASES.get(sheet_key, [])}
    rules = PLATFORM_SHEET_KEYWORDS.get(sheet_key) or {}
    must = [_norm_sheet(x) for x in rules.get("must", [])]
    reject = [_norm_sheet(x) for x in rules.get("reject", [])]
    hits = []
    for name in sheetnames:
        if _is_pivot_sheet(name):
            continue
        n = _norm_sheet(name)
        if any(token and token in n for token in reject):
            continue
        if n in aliases or any(token in n for token in must):
            hits.append(name)
    return hits


def resolve_platform_sheet(sheetnames: list[str], sheet_key: str) -> str | None:
    hits = platform_sheet_candidates(sheetnames, sheet_key)
    if not hits:
        return None
    if len(hits) > 1:
        detailed = [h for h in hits if "明细" in str(h)]
        if len(detailed) == 1:
            return detailed[0]
        raise SystemExit(
            f"Platform business {sheet_key} matched {len(hits)} detail sheets: {hits}. "
            "Add an exact alias so exactly one sheet matches; refusing to pick one."
        )
    return hits[0]


def resolve_col(df: pd.DataFrame, *names: str) -> str:
    by_norm = {_norm_col(c): c for c in df.columns}
    for name in names:
        hit = by_norm.get(_norm_col(name))
        if hit:
            return hit
    raise SystemExit(
        f"Missing column(s) {list(names)}; available={list(df.columns)}"
    )


def col_series(df: pd.DataFrame, *names: str) -> pd.Series:
    return df[resolve_col(df, *names)]


def optional_col(df: pd.DataFrame, *names: str) -> pd.Series:
    by_norm = {_norm_col(c): c for c in df.columns}
    for name in names:
        hit = by_norm.get(_norm_col(name))
        if hit:
            return df[hit]
    return pd.Series(pd.NA, index=df.index)


def pick_system_hub(df: pd.DataFrame) -> pd.Series:
    """System Hub: first_ascan_hub, else origin_poe. Blank if both empty."""
    ascan_hub = norm_text(optional_col(df, "first_ascan_hub"))
    origin = norm_text(optional_col(df, "origin_poe"))
    return ascan_hub.where(ascan_hub.notna(), origin)


def pick_master_hub(ascan_s, plt_s, sys_s, ags_s):
    """Final Hub: System first_ascan_hub > Platform > System origin_poe > AGS."""
    out = pd.Series(pd.NA, index=plt_s.index, dtype="object")
    src = pd.Series(pd.NA, index=plt_s.index, dtype="object")
    for series, name in (
        (ascan_s, "System"),
        (plt_s, "Platform"),
        (sys_s, "System"),
        (ags_s, "AGS"),
    ):
        take = series.notna() & out.isna()
        out = out.mask(take, series)
        src = src.mask(take, name)
    return out, src


def classify_csv(path: Path) -> str | None:
    """System handover_pickup extracts may arrive as CSV instead of xlsx."""
    try:
        header = pd.read_csv(path, nrows=0)
    except Exception:
        return None
    cols = {str(c) for c in header.columns}
    if "pickup_poe_51005" in cols and "handover_time_local" in cols:
        return "System"
    return None


def classify_workbook(path: Path) -> str | None:
    if path.suffix.lower() == ".csv":
        return classify_csv(path)
    first_headers = read_sheet_headers(path)
    sheets = list(first_headers)
    if "Tracking Details" in sheets and "MAWB Details" in sheets:
        return "AGS"
    cols = set()
    for h in first_headers.values():
        cols.update(str(c) for c in h)
    if "pickup_poe_51005" in cols and "handover_time_local" in cols:
        return "System"
    if _is_platform_workbook(path, sheets):
        return "Platform"
    return None


def find_input_files() -> dict:
    found = {}
    print("=== Source file discovery ===", flush=True)
    candidates = sorted(list(ROOT.glob("*.xlsx")) + list(ROOT.glob("*.csv")))
    if not candidates:
        print("  (no .xlsx/.csv files in this folder)", flush=True)
    for p in candidates:
        if _skip_source_candidate(p):
            print(f"  {p.name} -> skipped", flush=True)
            continue
        kind = classify_workbook(p)
        print(f"  {p.name} -> {kind or 'unrecognized'}", flush=True)
        if kind:
            if kind in found:
                raise SystemExit(f"Duplicate {kind} file: {found[kind].name} and {p.name}")
            found[kind] = p
    print("=== Bound sources ===", flush=True)
    for kind in ("AGS", "System", "Platform"):
        path = found.get(kind)
        print(f"  {kind}: {path.name if path else '(missing)'}", flush=True)
    missing = {"AGS", "System", "Platform"} - set(found)
    if not found:
        raise SystemExit("No source files found (need at least one of AGS / System / Platform).")
    if missing:
        print(
            f"WARNING: Missing source files this run: {sorted(missing)}. "
            "Proceeding with available sources; missing-source parcels will not enter today's three-source master.",
            flush=True,
        )
    return found


def empty_std() -> pd.DataFrame:
    df = pd.DataFrame(columns=[
        "TNO", "Source", "File名称", "业务名称", "MAWB", "Customer", "Hub",
        "交接时间", "交接时间类型", "Ascan时间", "platform_sheet",
        "mawb_mapping_conflict_flag",
    ])
    df["交接时间"] = pd.Series(dtype="datetime64[ns]")
    df["Ascan时间"] = pd.Series(dtype="datetime64[ns]")
    return df


def process_ags(path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    print("  AGS Tracking Details", flush=True)
    tracking = read_excel_fast(path, sheet_name="Tracking Details")
    mawb = read_excel_fast(path, sheet_name="MAWB Details")
    invalid = tracking.loc[is_blank_series(clean_tn_series(tracking["tracking_number"]))].copy()
    tracking["TNO"] = clean_tn_series(tracking["tracking_number"])
    tracking = tracking[tracking["TNO"].notna()].copy()
    before = len(tracking)
    tracking["_nn"] = (~is_blank_series(tracking["mawb"])).astype(int) + (~is_blank_series(tracking["handoverTime_local"])).astype(int)
    tracking = tracking.sort_values(["TNO", "_nn"], ascending=[True, False])
    dup_n = int(tracking["TNO"].duplicated().sum())
    tracking = tracking.drop_duplicates("TNO", keep="first").drop(columns=["_nn"])
    print(f"    tracking rows {before} -> unique TNO {len(tracking)} (dropped_dup_rows={dup_n})", flush=True)

    tracking["mawb_key"] = clean_mawb_series(tracking["mawb"])
    mawb["mawb_key"] = clean_mawb_series(mawb["MAWB"])
    mawb_n = mawb.groupby("mawb_key", dropna=True).agg(
        n_cust=("Customer", lambda s: s.dropna().astype(str).str.strip().nunique()),
        n_hub=("Hub", lambda s: s.dropna().astype(str).str.strip().nunique()),
        Customer=("Customer", "first"),
        Hub=("Hub", "first"),
    )
    conflict_keys = set(mawb_n.index[(mawb_n["n_cust"] > 1) | (mawb_n["n_hub"] > 1)])
    mawb_ok = mawb_n.loc[~mawb_n.index.isin(conflict_keys), ["Customer", "Hub"]]
    tracking = tracking.merge(mawb_ok, left_on="mawb_key", right_index=True, how="left")
    tracking["mawb_mapping_conflict_flag"] = np.where(
        tracking["mawb_key"].isin(conflict_keys), "Yes", "No"
    )
    tracking.loc[tracking["mawb_mapping_conflict_flag"] == "Yes", ["Customer", "Hub"]] = pd.NA

    mawb_conflicts = []
    for k in conflict_keys:
        sub = mawb[mawb["mawb_key"] == k]
        mawb_conflicts.append({
            "TNO": None,
            "field_name": "MAWB_mapping",
            "ags_raw": f"MAWB={k}; customers={list(sub['Customer'].unique())}; hubs={list(sub['Hub'].unique())}",
            "system_raw": None,
            "platform_raw": None,
            "chosen_value": None,
            "chosen_reason": "mawb_mapping_conflict_left_blank",
            "needs_manual_review": "Yes",
            "primary_source": "AGS",
            "cohort_date": None,
        })
    stem = path.stem
    file_label = stem if stem.startswith("AGS_to_SPX") else f"AGS_to_SPX_24H_Ascan_Report_{path.stem}"

    std = pd.DataFrame({
        "TNO": tracking["TNO"],
        "Source": "AGS",
        "File名称": file_label,
        "业务名称": "AGS",
        "MAWB": tracking["mawb"].where(~is_blank_series(tracking["mawb"]), None),
        "Customer": tracking["Customer"],
        "Hub": tracking["Hub"],
        "交接时间": parse_dt_series(tracking["handoverTime_local"]),
        "交接时间类型": "Handover Time",
        "Ascan时间": parse_dt_series(tracking["ascanTime"]),
        "platform_sheet": pd.NA,
        "mawb_mapping_conflict_flag": tracking["mawb_mapping_conflict_flag"],
        "ags_mawb_hub": tracking["Hub"],
        "ags_tracking_origin_hub": tracking["origin_hub"] if "origin_hub" in tracking.columns else pd.NA,
        "ags_destination_hub": tracking["destination_hub"] if "destination_hub" in tracking.columns else pd.NA,
        "ags_destination_facility": tracking["destination_facility"] if "destination_facility" in tracking.columns else pd.NA,
        "ags_handover_time": parse_dt_series(tracking["handoverTime_local"]),
        "ags_ascan_time": parse_dt_series(tracking["ascanTime"]),
    })
    for col in ("ags_tracking_origin_hub", "ags_destination_hub", "ags_destination_facility", "ags_mawb_hub"):
        if col in std.columns:
            std[col] = norm_text(std[col])
    return std, invalid, pd.DataFrame(mawb_conflicts)


def process_system(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    print("  System", flush=True)
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=object, low_memory=False)
    else:
        df = read_excel_fast(path, sheet_name=0)
    df["TNO"] = clean_tn_series(df["tracking_number"])
    invalid = df[df["TNO"].isna()].copy()
    df = df[df["TNO"].notna()].copy()
    df = df.drop_duplicates("TNO", keep="first")
    hub = pick_system_hub(df)
    ascan_hub = norm_text(optional_col(df, "first_ascan_hub"))
    n_ascan = int(ascan_hub.notna().sum())
    n_origin = int((ascan_hub.isna() & hub.notna()).sum())
    n_drop = int(hub.isna().sum())
    print(
        f"    hub first_ascan_hub={n_ascan:,} origin_poe_fallback={n_origin:,} "
        f"dropped_both_empty={n_drop:,}",
        flush=True,
    )
    keep = hub.notna()
    df = df.loc[keep].copy()
    hub = hub.loc[keep]
    ascan_hub = ascan_hub.loc[keep]
    pickup = parse_dt_series(df["pickup_poe_51005"])
    handover = parse_dt_series(df["handover_time_local"])
    both = pickup.notna() & handover.notna()
    is_tt = df["customer_name"].astype("string").str.contains(TIKTOK_RE, na=False)
    start = pickup.copy()
    start_type = pd.Series(pd.NA, index=df.index, dtype="object")
    start_type = start_type.mask(pickup.notna(), "Pickup Time")
    start = start.mask(handover.notna() & pickup.isna(), handover)
    start_type = start_type.mask(handover.notna() & pickup.isna(), "Handover Time")
    # both filled: TikTok -> Handover, else Pickup
    start = start.mask(both & is_tt, handover)
    start_type = start_type.mask(both & is_tt, "Handover Time")
    start = start.mask(both & ~is_tt, pickup)
    start_type = start_type.mask(both & ~is_tt, "Pickup Time")

    cust = df["customer_name"].astype("string").str.strip()
    biz = cust.copy()
    std = pd.DataFrame({
        "TNO": df["TNO"],
        "Source": "System",
        "File名称": path.stem,
        "业务名称": biz,
        "MAWB": df["mawb"].where(~is_blank_series(df["mawb"]), None),
        "Customer": cust,
        "Hub": hub,
        "system_first_ascan_hub": ascan_hub,
        "交接时间": start,
        "交接时间类型": start_type,
        "Ascan时间": parse_dt_series(df["first_ascan_poe"]),
        "platform_sheet": pd.NA,
        "mawb_mapping_conflict_flag": "No",
        "_both_pickup_handover": both.astype(int),
    })
    return std, invalid


def _sheet_tracking_numbers(path: Path, sheet: str) -> set[str]:
    df = read_excel_fast(
        path,
        sheet_name=sheet,
        usecols=lambda c: _norm_col(c) in TRACKING_HEADER_NAMES,
    )
    if df.shape[1] == 0:
        return set()
    series = df.iloc[:, 0].dropna().astype(str).str.strip()
    return set(series[series != ""])


def drop_covered_annotation_sheets(path: Path, sheetnames: list[str]) -> list[str]:
    """Skip 整箱未扫 tabs when every tracking number is already on another detail sheet."""
    annotations = [s for s in sheetnames if "整箱未扫" in str(s) and not _is_pivot_sheet(s)]
    if not annotations:
        return sheetnames
    kept = [s for s in sheetnames if s not in annotations]
    for name in annotations:
        n = _norm_sheet(name)
        siblings = []
        for other in kept:
            if _is_pivot_sheet(other):
                continue
            on = _norm_sheet(other)
            for rules in PLATFORM_SHEET_KEYWORDS.values():
                must = [_norm_sheet(x) for x in rules.get("must", [])]
                reject = [_norm_sheet(x) for x in rules.get("reject", [])]
                if any(token and token in n for token in reject):
                    continue
                if any(token and token in n and token in on for token in must):
                    siblings.append(other)
                    break
        if not siblings:
            raise SystemExit(
                f"Annotation sheet {name!r} has no sibling detail sheet to check coverage against."
            )
        ann_tn = _sheet_tracking_numbers(path, name)
        covered_by = None
        for sib in siblings:
            if ann_tn <= _sheet_tracking_numbers(path, sib):
                covered_by = sib
                break
        if covered_by is None:
            raise SystemExit(
                f"Annotation sheet {name!r} has tracking numbers that are not in {siblings}. "
                "Refusing to drop them."
            )
        print(
            f"    skip annotation sheet {name!r}: {len(ann_tn)} TNO already in {covered_by!r}",
            flush=True,
        )
    return kept


def process_platform(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    print("  Platform sheets", flush=True)
    headers = read_sheet_headers(path)
    sheetnames = list(headers)
    print(f"    workbook sheets: {sheetnames}", flush=True)
    sheetnames = drop_covered_annotation_sheets(path, sheetnames)
    specs = [
        {
            "sheet_key": "TT-GS",
            "file_label": "平台客户需ASCAN单号明细 - TT-GS",
            "biz": "TT GS",
            "tn": ["追踪号", "Tracking Number"],
            "mawb": ["主运单号", "MAWB"],
            "customer_const": "TT GS",
            "hub": ["卡转线路", "Origin POE", "Origin Hub", "实际入网口岸"],
            "start": ["交接当地时间"],
            "start_type": "Handover Time",
            "ascan": ["首次扫描当地时间"],
            "platform_sheet": "TT-GS",
        },
        {
            "sheet_key": "FBT",
            "file_label": "平台客户需ASCAN单号明细 - FBT",
            "biz": "TT FBT",
            "tn": ["Tracking Number", "追踪号"],
            "mawb": [],
            "customer_const": "TT FBT",
            "hub": ["Origin Hub", "Origin POE"],
            "start": ["Manifest Time", "date_manifested"],
            "start_type": "Manifest Time",
            "ascan": ["First Arrival Scan", "First Scan Date"],
            "platform_sheet": "FBT",
        },
        {
            "sheet_key": "Winit",
            "file_label": "平台客户需ASCAN单号明细 - Winit",
            "biz": "Winit",
            "tn": ["Tracking Number", "追踪号"],
            "mawb": [],
            "customer_const": "Winit",
            "hub": ["Origin Hub", "Origin POE"],
            "start": ["date_manifested", "Manifest Time"],
            "start_type": "Manifest Time",
            "ascan": ["First Scan Date"],
            "platform_sheet": "Winit",
        },
        {
            "sheet_key": "AliExpress",
            "file_label": "平台客户需ASCAN单号明细 - AliExpress",
            "biz": "AliExpress",
            "tn": ["跟踪号", "Tracking Number", "追踪号"],
            "mawb": ["提单号", "主运单号"],
            "customer_const": "AliExpress",
            "hub": ["Origin POE", "Origin Hub"],
            "start": ["清关交付-SPX收货 时间", "清关交付-SPX收货时间"],
            "start_type": "Handover Time",
            "ascan": ["First Scan Arrival Scan", "Ascan收货时间"],
            "platform_sheet": "AliExpress",
        },
    ]
    assigned: dict[str, str] = {}
    for spec in specs:
        actual = resolve_platform_sheet(sheetnames, spec["sheet_key"])
        if actual is None:
            continue
        if actual in assigned:
            raise SystemExit(
                f"Sheet {actual!r} matched both {assigned[actual]} and {spec['sheet_key']}; "
                "tighten PLATFORM_SHEET_ALIASES."
            )
        assigned[actual] = spec["sheet_key"]
    report_platform_sheets(path, sheetnames, assigned, headers)
    sheet_by_key = {key: sheet for sheet, key in assigned.items()}

    parts = []
    invalid_parts = []
    for spec in specs:
        actual = sheet_by_key.get(spec["sheet_key"])
        if actual is None:
            print(f"    {spec['sheet_key']}: 业务缺席，本工作簿无对应明细表", flush=True)
            continue
        df = read_excel_fast(path, sheet_name=actual)
        try:
            tn_col = resolve_col(df, *spec["tn"])
            hub_col = resolve_col(df, *spec["hub"])
            start_col = resolve_col(df, *spec["start"])
            ascan_col = resolve_col(df, *spec["ascan"])
            mawb_col = resolve_col(df, *spec["mawb"]) if spec["mawb"] else None
        except SystemExit as exc:
            raise SystemExit(
                f"Platform {spec['sheet_key']} sheet {actual!r} 已识别但字段解析失败: {exc}"
            ) from exc
        print(
            f"    {spec['sheet_key']} <- {actual!r}; "
            f"tn={tn_col!r} mawb={mawb_col!r} hub={hub_col!r} "
            f"start={start_col!r} ascan={ascan_col!r}",
            flush=True,
        )
        raw_n = len(df)
        df["TNO"] = clean_tn_series(df[tn_col])
        empty_n = int(df["TNO"].isna().sum())
        invalid_parts.append(df[df["TNO"].isna()].copy())
        df = df[df["TNO"].notna()]
        valid_n = len(df)
        df = df.drop_duplicates("TNO", keep="first")
        print(
            f"      rows={raw_n:,} empty_TNO={empty_n:,} valid={valid_n:,} "
            f"unique={len(df):,} in_sheet_dup={valid_n - len(df):,}",
            flush=True,
        )
        if mawb_col:
            mawb = df[mawb_col].where(~is_blank_series(df[mawb_col]), None)
        else:
            mawb = None
        std = pd.DataFrame({
            "TNO": df["TNO"],
            "Source": "Platform",
            "File名称": spec["file_label"],
            "业务名称": spec["biz"],
            "MAWB": mawb,
            "Customer": spec["customer_const"],
            "Hub": df[hub_col].astype("string").str.strip(),
            "交接时间": parse_dt_series(df[start_col]),
            "交接时间类型": spec["start_type"],
            "Ascan时间": parse_dt_series(df[ascan_col]),
            "platform_sheet": spec["platform_sheet"],
            "mawb_mapping_conflict_flag": "No",
        })
        parts.append(std)
    if not parts:
        raise SystemExit(
            f"Platform file {path.name} has no recognizable detail sheets. "
            f"Found sheets: {sheetnames}"
        )
    invalid = pd.concat(invalid_parts, ignore_index=True) if invalid_parts else pd.DataFrame()
    merged = pd.concat(parts, ignore_index=True)
    cross = merged.loc[merged["TNO"].duplicated(keep=False), ["TNO", "platform_sheet"]]
    if len(cross):
        pairs = cross.groupby("TNO")["platform_sheet"].apply(lambda s: ",".join(sorted(set(s))))
        raise SystemExit(
            f"Platform cross-sheet duplicate TNO: {pairs.nunique():,} distinct combos, "
            f"{cross['TNO'].nunique():,} TNO. Sample: {pairs.head(10).to_dict()}. "
            "One TNO must not appear in two business detail sheets."
        )
    print(f"    platform combined unique TNO = {merged['TNO'].nunique():,}", flush=True)
    return merged, invalid


def norm_text(s: pd.Series) -> pd.Series:
    t = s.astype("string").str.strip()
    t = t.replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA, "None": pd.NA})
    return t


def norm_text_upper(s: pd.Series) -> pd.Series:
    return norm_text(s).str.upper()


def is_tiktok_alias(a, b) -> bool:
    aa = None if a is None or (isinstance(a, float) and pd.isna(a)) else str(a).strip().upper()
    bb = None if b is None or (isinstance(b, float) and pd.isna(b)) else str(b).strip().upper()
    if not aa or not bb:
        return False
    def key(x):
        if x in TIKTOK_CUSTOMERS or "TIKTOK" in x:
            return "TIKTOK"
        return x
    return key(aa) == key(bb) == "TIKTOK"


EMPTY_TEXT = {"", "nan", "none", "<na>", "nat", "null", "<n/a>"}


def _is_missing_scalar(value) -> bool:
    if value is None:
        return True
    try:
        # pd.isna(pd.NA) is True; do not use `if value`
        return bool(pd.isna(value))
    except (ValueError, TypeError):
        return False


def _as_clean_text(value) -> str | None:
    if isinstance(value, (pd.Series, pd.DataFrame)):
        return None
    if _is_missing_scalar(value):
        return None
    text = str(value).strip()
    if text.lower() in EMPTY_TEXT:
        return None
    return text


def _flatten_values(values):
    if isinstance(values, pd.Series):
        values = values.tolist()
    elif not isinstance(values, (list, tuple, set)):
        values = [values]
    out = []
    for value in values:
        if isinstance(value, pd.Series):
            out.extend(value.tolist())
        elif isinstance(value, (list, tuple, set)):
            out.extend(list(value))
        else:
            out.append(value)
    return out


def join_nonempty_values(values, sep=";"):
    result = []
    for value in _flatten_values(values):
        text = _as_clean_text(value)
        if text is None:
            continue
        if text not in result:
            result.append(text)
    return sep.join(result) if result else None


def join_nonempty(series_list, sep=";"):
    """Join aligned Series row-wise. Never boolean-test pd.NA."""
    cleaned = []
    for s in series_list:
        t = s.astype("string").fillna("")
        t = t.str.strip()
        t = t.mask(t.str.lower().isin(sorted(EMPTY_TEXT)), "")
        cleaned.append(t.astype(str))
    acc = cleaned[0]
    for t in cleaned[1:]:
        empty_acc = acc.eq("")
        empty_t = t.eq("")
        same = acc.eq(t)
        only_t = empty_acc & ~empty_t
        both = ~empty_acc & ~empty_t & ~same
        acc = acc.mask(only_t, t)
        acc = acc.mask(both, acc + sep + t)
    return acc.mask(acc.eq(""), None)


def _norm_customer_alias_key(value) -> str | None:
    text = _as_clean_text(value)
    if text is None:
        return None
    key = re.sub(r"\s+", " ", text).upper()
    return key.replace(" ,", ",").replace(", ", ",")


_CUSTOMER_ALIAS_LOOKUP = {
    k: v for k, v in ((_norm_customer_alias_key(src), dst) for src, dst in CUSTOMER_ALIASES.items()) if k
}


def is_tiktok_customer_series(s: pd.Series) -> pd.Series:
    """True for TIKTOK / TikTok Inc / any TikTok case variant. Not TT GS or TT FBT."""
    return s.astype("string").str.contains(TIKTOK_RE, na=False)


def is_cainiao_customer_series(s: pd.Series) -> pd.Series:
    t = s.astype("string")
    return t.str.contains(r"HANGZHOU CAINIAO SUPPLY CHAIN MANAGEMENT", case=False, na=False)


def coalesce_customer(*series_list: pd.Series) -> pd.Series:
    out = series_list[0].astype("object")
    for s in series_list[1:]:
        out = out.mask(out.isna() & s.notna(), s)
    return out


def standardize_ags_customer(picked: pd.Series, ags_cust: pd.Series) -> tuple[pd.Series, pd.Series, pd.Series, pd.Series]:
    """Keep AGS original in customer_raw. Display TT GS / AliExpress when AGS matches."""
    customer_raw = coalesce_customer(ags_cust, picked)
    std = picked.astype("object")
    tt = is_tiktok_customer_series(ags_cust)
    cainiao = is_cainiao_customer_series(ags_cust)
    std = std.mask(tt, "TT GS")
    std = std.mask(cainiao, "AliExpress")
    return customer_raw, std, tt, cainiao


def pick_nonempty(plt_s, sys_s, ags_s, is_dt=False):
    """Field-level fill: Platform > System > AGS. Empty values never overwrite."""
    if is_dt:
        out = pd.Series(pd.NaT, index=plt_s.index)
    else:
        out = pd.Series(pd.NA, index=plt_s.index, dtype="object")
    src = pd.Series(pd.NA, index=plt_s.index, dtype="object")
    for series, name in ((plt_s, "Platform"), (sys_s, "System"), (ags_s, "AGS")):
        take = series.notna() & out.isna()
        out = out.mask(take, series)
        src = src.mask(take, name)
    return out, src


def _dest_hub_set(series: pd.Series) -> set[str]:
    out = set()
    for v in series:
        text = _as_clean_text(v)
        if text:
            out.add(text.upper())
    return out


def tag_ags_inter_hub(ags: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Tag AGS TNOs using MAWB Hub + destination_hub only. Not unified_master.Hub."""
    print("  AGS inter-hub transfer tags", flush=True)
    df = ags.copy()
    mawb_hub = norm_text(df["ags_mawb_hub"]).str.upper()
    dest = norm_text(df["ags_destination_hub"]).str.upper()
    mawb_key = clean_mawb_series(df["MAWB"])
    df["_hub"] = mawb_hub
    df["_dest"] = dest
    df["_mawb"] = mawb_key

    summary_rows = []
    tag_by_key = {}
    dests_by_key = {}
    grouped = df.groupby(["_hub", "_mawb"], dropna=False, sort=False)
    for (hub, mawb), sub in grouped:
        dests = _dest_hub_set(sub["_dest"])
        dest_label = ",".join(sorted(dests)) if dests else ""
        hub_ok = isinstance(hub, str) and hub in ("LAX", "EWR")
        mawb_ok = pd.notna(mawb) and str(mawb).strip() not in ("", "nan", "<NA>")
        tag = None
        if hub_ok and mawb_ok:
            if hub == "LAX" and dests and dests <= LAX_TEXAS_DESTS:
                tag = "LAX-DFW"
            elif hub == "EWR" and EWR_DOWNSTREAM_DESTS <= dests:
                tag = "EWR-ORD"
        key = (hub, mawb)
        dests_by_key[key] = dest_label
        if tag:
            tag_by_key[key] = tag
            summary_rows.append({
                "ags_mawb_hub": hub,
                "mawb": mawb,
                "ags_destination_hubs": dest_label,
                "跨Hub转运": tag,
                "tno_count": int(sub["TNO"].nunique()),
            })
    df["ags_destination_hubs"] = [
        dests_by_key.get((h, m), "") for h, m in zip(df["_hub"], df["_mawb"])
    ]
    df["跨Hub转运"] = [
        tag_by_key.get((h, m)) for h, m in zip(df["_hub"], df["_mawb"])
    ]
    df = df.drop(columns=["_hub", "_dest", "_mawb"])
    summary = pd.DataFrame(summary_rows)
    if summary.empty:
        print("    tagged MAWBs=0 TNO=0", flush=True)
    else:
        for tag, sub in summary.groupby("跨Hub转运"):
            print(
                f"    {tag}: MAWB={len(sub):,} TNO={int(sub['tno_count'].sum()):,}",
                flush=True,
            )
        print("    tagged MAWB list:", flush=True)
        for _, r in summary.sort_values(["跨Hub转运", "mawb"]).iterrows():
            print(
                f"      {r['跨Hub转运']} hub={r['ags_mawb_hub']} mawb={r['mawb']} dests={r['ags_destination_hubs']} tno={int(r['tno_count'])}",
                flush=True,
            )
    return df, summary


def attach_ags_inter_hub(master: pd.DataFrame, ags: pd.DataFrame) -> pd.DataFrame:
    keep = ["TNO"] + [c for c in AGS_EXTRA_COLS if c in ags.columns]
    extra = ags[keep].drop_duplicates("TNO").copy()
    drop_existing = [c for c in extra.columns if c != "TNO" and c in master.columns]
    master = master.drop(columns=drop_existing)
    if master.index.name == "TNO" or "TNO" in getattr(master.index, "names", []):
        master = master.reset_index(drop=True)
    extra["TNO"] = extra["TNO"].astype(str)
    master["TNO"] = master["TNO"].astype(str)
    master = master.merge(extra, on="TNO", how="left")
    if "跨Hub转运" not in master.columns:
        master["跨Hub转运"] = pd.NA
    allowed = master["跨Hub转运"].astype("string").isin(["LAX-DFW", "EWR-ORD"])
    master["跨Hub转运"] = master["跨Hub转运"].where(allowed, pd.NA)
    return master


def merge_sources(ags: pd.DataFrame, sys_df: pd.DataFrame, plt: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    print("  merge by TNO", flush=True)
    a = ags.set_index("TNO", drop=False)
    s = sys_df.set_index("TNO", drop=False)
    p = plt.set_index("TNO", drop=False)
    all_tns = a.index.union(s.index).union(p.index)
    idx = pd.Index(all_tns, name="TNO")

    def align(df, col, fill=None):
        if col not in df.columns:
            return pd.Series(pd.NA, index=idx)
        return df[col].reindex(idx)

    has_ags = idx.isin(a.index)
    has_sys = idx.isin(s.index)
    has_plt = idx.isin(p.index)
    plt_sheet = align(p, "platform_sheet")
    is_gs = has_plt & (plt_sheet == "TT-GS")
    is_ali = has_plt & (plt_sheet == "AliExpress")
    tiktok_gs_with_sys = has_sys & is_gs
    ali_with_ags = is_ali & has_ags

    primary = pd.Series(pd.NA, index=idx, dtype="object")
    primary = primary.mask(has_plt, "Platform")
    primary = primary.mask(primary.isna() & has_sys, "System")
    primary = primary.mask(primary.isna() & has_ags, "AGS")

    cust_a = norm_text(align(a, "Customer"))
    cust_s = norm_text(align(s, "Customer"))
    cust_p = norm_text(align(p, "Customer"))
    hub_a = norm_text(align(a, "Hub"))
    hub_s = norm_text(align(s, "Hub"))
    hub_p = norm_text(align(p, "Hub"))
    mawb_a = norm_text(align(a, "MAWB"))
    mawb_s = norm_text(align(s, "MAWB"))
    mawb_p = norm_text(align(p, "MAWB"))
    biz_a = norm_text(align(a, "业务名称"))
    biz_s = norm_text(align(s, "业务名称"))
    biz_p = norm_text(align(p, "业务名称"))
    start_a = parse_dt_series(align(a, "交接时间"))
    start_s = parse_dt_series(align(s, "交接时间"))
    start_p = parse_dt_series(align(p, "交接时间"))
    type_a = align(a, "交接时间类型")
    type_s = align(s, "交接时间类型")
    type_p = align(p, "交接时间类型")
    ascan_a = parse_dt_series(align(a, "Ascan时间"))
    ascan_s = parse_dt_series(align(s, "Ascan时间"))
    ascan_p = parse_dt_series(align(p, "Ascan时间"))

    biz, _biz_src = pick_nonempty(biz_p, biz_s, biz_a)
    picked_cust, cust_src = pick_nonempty(cust_p, cust_s, cust_a)
    _ags_raw, cust, ags_tt, ags_cainiao = standardize_ags_customer(picked_cust, cust_a)
    customer_raw = join_nonempty([cust_a, cust_s, cust_p], sep=";")
    mapped_alias = ags_cainiao
    biz = biz.mask(ags_tt, "TT GS")
    biz = biz.mask(ags_cainiao, "AliExpress")
    ascan_hub_s = norm_text(align(s, "system_first_ascan_hub"))
    hub, hub_src = pick_master_hub(ascan_hub_s, hub_p, hub_s, hub_a)
    mawb, mawb_src = pick_nonempty(mawb_p, mawb_s, mawb_a)
    start, start_src = pick_nonempty(start_p, start_s, start_a, is_dt=True)
    ascan, ascan_src = pick_nonempty(ascan_p, ascan_s, ascan_a, is_dt=True)

    start_type = pd.Series(pd.NA, index=idx, dtype="object")
    start_type = start_type.mask(start_src.eq("Platform"), type_p)
    start_type = start_type.mask(start_src.eq("System"), type_s)
    start_type = start_type.mask(start_src.eq("AGS"), type_a)
    type_fill, _ = pick_nonempty(type_p, type_s, type_a)
    start_type = start_type.mask(start_type.isna() & type_fill.notna(), type_fill)
    start_type = display_start_type(start_type)

    filled_from_sys_ags = has_plt & (
        cust_src.isin(["System", "AGS"])
        | hub_src.isin(["System", "AGS"])
        | mawb_src.isin(["System", "AGS"])
        | start_src.isin(["System", "AGS"])
        | ascan_src.isin(["System", "AGS"])
    )
    note_parts = pd.Series("", index=idx, dtype="object")
    note_parts = note_parts.mask(has_plt, note_parts + "platform_preferred;")
    note_parts = note_parts.mask(filled_from_sys_ags, note_parts + "system_ags_fill_empty_platform_fields;")
    note_parts = note_parts.mask(tiktok_gs_with_sys, note_parts + "tiktok_gs_customer_tt_gs;")
    note_parts = note_parts.mask(ali_with_ags, note_parts + "ali_platform_customer_ags_in_audit;")
    hub_ascan_over_plt = (
        ascan_hub_s.notna() & hub_p.notna()
        & (ascan_hub_s.str.upper() != hub_p.str.upper())
        & (hub_src == "System")
    )
    note_parts = note_parts.mask(hub_ascan_over_plt, note_parts + "hub_system_first_ascan_hub_priority;")
    hub_conflict_plt = hub_p.notna() & hub_s.notna() & (hub_p.str.upper() != hub_s.str.upper()) & (hub_src == "Platform")
    note_parts = note_parts.mask(hub_conflict_plt, note_parts + "hub_platform_priority_over_system;")
    note_parts = note_parts.mask(mapped_alias, note_parts + "cainiao_mapped_to_aliexpress;")
    note_parts = note_parts.mask(ags_tt, note_parts + "ags_tiktok_standardized_to_tt_gs;")
    merge_note = note_parts.str.strip(";").replace({"": None})

    source_files = join_nonempty([
        align(a, "File名称"),
        align(s, "File名称"),
        align(p, "File名称"),
    ], sep=";")
    source = join_nonempty([
        pd.Series(np.where(has_ags, "AGS", None), index=idx),
        pd.Series(np.where(has_sys, "System", None), index=idx),
        pd.Series(np.where(has_plt, "Platform", None), index=idx),
    ], sep=";")

    cohort = start.dt.strftime("%Y-%m-%d")
    first_ascan = ascan
    td = (ascan - start) / pd.Timedelta(hours=1)
    within = pd.Series(pd.NA, index=idx, dtype="object")
    both_t = start.notna() & ascan.notna()
    within = within.mask(both_t, np.where(td <= 24, "Yes", "No"))
    monitor = pd.Series("Yes", index=idx, dtype="object")
    monitor = monitor.mask(both_t & (td <= 24), "No")

    mawb_flag = align(a, "mawb_mapping_conflict_flag").fillna("No")

    master = pd.DataFrame({
        "Source": source,
        "File名称": source_files,
        "业务名称": biz,
        "TNO": idx.astype(str),
        "MAWB": mawb,
        "Customer": cust,
        "CUSTOMER": cust,
        "customer_raw": customer_raw,
        "customer_standardized": cust,
        "Hub": hub,
        "交接时间": start,
        "交接时间类型": start_type,
        "Ascan时间": ascan,
        "start_time": start,
        "start_time_type": start_type,
        "cohort_date": cohort,
        "first_ascan_time": first_ascan,
        "time_diff_hours": td.round(4),
        "within_24h": within,
        "monitor_flag": monitor,
        "primary_source": primary,
        "merge_note": merge_note,
        "source_files": source_files,
        "source_count": has_ags.astype(int) + has_sys.astype(int) + has_plt.astype(int),
        "has_ags": np.where(has_ags, "Yes", "No"),
        "has_system": np.where(has_sys, "Yes", "No"),
        "has_platform": np.where(has_plt, "Yes", "No"),
        "customer_source": cust_src,
        "hub_source": hub_src,
        "start_time_source": start_src,
        "ascan_source": ascan_src,
        "mawb_source": mawb_src,
        "mawb_mapping_conflict_flag": mawb_flag,
        "platform_sheet": plt_sheet,
        "_cust_a": cust_a, "_cust_s": cust_s, "_cust_p": cust_p,
        "_hub_a": hub_a, "_hub_s": hub_s, "_hub_p": hub_p,
        "_mawb_a": mawb_a, "_mawb_s": mawb_s, "_mawb_p": mawb_p,
        "_start_a": start_a, "_start_s": start_s, "_start_p": start_p,
        "_ascan_a": ascan_a, "_ascan_s": ascan_s, "_ascan_p": ascan_p,
        "_type_a": type_a, "_type_s": type_s, "_type_p": type_p,
        "_ali_with_ags": ali_with_ags,
        "_tiktok_gs_sys": tiktok_gs_with_sys,
    })
    master, n_fbt_iah = remap_fbt_dfw_hub(master)
    if n_fbt_iah:
        print(f"  TT FBT Hub DFW -> IAH: {n_fbt_iah:,}", flush=True)

    print("  build conflict audit", flush=True)

    def dump_series(s: pd.Series) -> pd.Series:
        if pd.api.types.is_datetime64_any_dtype(s):
            return s.dt.strftime("%Y-%m-%d %H:%M:%S")
        return s.astype("string")

    def field_audit(field, aa, ss, pp, chosen, reason, manual, is_dt=False):
        filled = (aa.notna().astype(int) + ss.notna().astype(int) + pp.notna().astype(int)) >= 2
        if is_dt:
            ca, cs, cp = parse_dt_series(aa), parse_dt_series(ss), parse_dt_series(pp)
            differ = (
                (aa.notna() & ss.notna() & (ca != cs))
                | (aa.notna() & pp.notna() & (ca != cp))
                | (ss.notna() & pp.notna() & (cs != cp))
            )
        else:
            ca, cs, cp = norm_text_upper(aa), norm_text_upper(ss), norm_text_upper(pp)
            differ = (
                (aa.notna() & ss.notna() & (ca != cs))
                | (aa.notna() & pp.notna() & (ca != cp))
                | (ss.notna() & pp.notna() & (cs != cp))
            )
        mask = filled & differ
        if not mask.any():
            return pd.DataFrame()
        out = pd.DataFrame({
            "TNO": master.loc[mask, "TNO"],
            "field_name": field,
            "ags_raw": dump_series(aa.loc[mask]),
            "system_raw": dump_series(ss.loc[mask]),
            "platform_raw": dump_series(pp.loc[mask]),
            "chosen_value": dump_series(chosen.loc[mask]),
            "chosen_reason": reason.loc[mask] if isinstance(reason, pd.Series) else reason,
            "needs_manual_review": manual.loc[mask] if isinstance(manual, pd.Series) else manual,
            "primary_source": master.loc[mask, "primary_source"],
            "cohort_date": master.loc[mask, "cohort_date"],
        })
        return out

    cust_reason = pd.Series("platform_preferred", index=idx, dtype="object")
    cust_reason = cust_reason.mask(cust_src == "System", "system_fill_empty_platform")
    cust_reason = cust_reason.mask(cust_src == "AGS", "ags_fill_empty_platform")
    cust_reason = cust_reason.mask(tiktok_gs_with_sys, "platform_preferred_tiktok_gs")
    cust_reason = cust_reason.mask(ali_with_ags, "platform_preferred_aliexpress")
    cust_reason = cust_reason.mask(mapped_alias, "cainiao_mapped_to_aliexpress")
    cust_reason = cust_reason.mask(ags_tt, "ags_tiktok_standardized_to_tt_gs")
    alias_tt = tiktok_gs_with_sys | (
        cust_a.astype("string").str.contains("TIKTOK", case=False, na=False)
        & (
            cust_s.astype("string").str.contains("TIKTOK", case=False, na=False)
            | cust_p.astype("string").eq("TT GS").fillna(False)
        )
    )
    cust_manual = pd.Series("Yes", index=idx, dtype="object")
    cust_manual = cust_manual.mask(alias_tt, "No")
    cust_manual = cust_manual.mask(tiktok_gs_with_sys, "No")
    cust_manual = cust_manual.mask(ali_with_ags, "No")
    cust_manual = cust_manual.mask(mapped_alias, "No")
    cust_manual = cust_manual.mask(ags_tt, "No")

    hub_from_ascan = ascan_hub_s.notna() & (hub_src == "System")
    hub_reason = pd.Series("hub_conflict", index=idx, dtype="object")
    hub_reason = hub_reason.mask(hub_src == "Platform", "platform_hub_when_no_system_first_ascan_hub")
    hub_reason = hub_reason.mask(hub_src == "System", "system_fill_empty_platform")
    hub_reason = hub_reason.mask(hub_from_ascan, "system_first_ascan_hub_priority")
    hub_reason = hub_reason.mask(hub_src == "AGS", "ags_fill_empty_platform")
    start_reason = pd.Series("start_time_conflict", index=idx, dtype="object")
    start_reason = start_reason.mask(start_src == "Platform", "platform_preferred")
    start_reason = start_reason.mask(start_src == "System", "system_fill_empty_platform")
    start_reason = start_reason.mask(start_src == "AGS", "ags_fill_empty_platform")
    start_date_a = start_a.dt.strftime("%Y-%m-%d")
    start_date_s = start_s.dt.strftime("%Y-%m-%d")
    start_date_p = start_p.dt.strftime("%Y-%m-%d")
    start_date_conflict = (
        (start_a.notna() & start_s.notna() & (start_date_a != start_date_s))
        | (start_a.notna() & start_p.notna() & (start_date_a != start_date_p))
        | (start_s.notna() & start_p.notna() & (start_date_s != start_date_p))
    )
    start_reason = start_reason.mask(start_date_conflict, "start_date_conflict")
    start_manual = pd.Series("No", index=idx, dtype="object")
    start_manual = start_manual.mask(start_date_conflict, "Yes")
    ascan_reason = pd.Series("platform_preferred", index=idx, dtype="object")
    ascan_reason = ascan_reason.mask(ascan_src == "System", "system_fill_empty_platform")
    ascan_reason = ascan_reason.mask(ascan_src == "AGS", "ags_fill_empty_platform")

    audit_parts = [
        field_audit("Customer", cust_a, cust_s, cust_p, cust, cust_reason, cust_manual, False),
        field_audit("Hub", hub_a, hub_s, hub_p, hub, hub_reason, pd.Series(np.where(hub_from_ascan | (hub_src == "Platform"), "No", "Yes"), index=idx), False),
        field_audit("MAWB", mawb_a, mawb_s, mawb_p, mawb, pd.Series("mawb_value_conflict", index=idx), pd.Series("Yes", index=idx), False),
        field_audit("交接时间", start_a, start_s, start_p, start, start_reason, start_manual, True),
        field_audit("Ascan时间", ascan_a, ascan_s, ascan_p, ascan, ascan_reason, pd.Series("No", index=idx), True),
    ]
    audit_df = pd.concat([x for x in audit_parts if not x.empty], ignore_index=True) if any(len(x) for x in audit_parts) else pd.DataFrame()
    drop_cols = [c for c in master.columns if c.startswith("_")]
    master = master.drop(columns=drop_cols)
    return master, audit_df


def output_path_or_fallback(path: Path) -> Path:
    """If the target file is open in Excel, write to *_new instead of failing."""
    try:
        with open(path, "ab"):
            pass
        return path
    except PermissionError:
        alt = path.with_name(path.stem + "_new" + path.suffix)
        print(f"  WARN {path.name} is locked, writing {alt.name}", flush=True)
        return alt


def write_display_files(display_dt: pd.DataFrame, biz_day: str, full_excel: bool = False):
    """Write the 18-column operational master as dated parquet (plus csv/xlsx on --full-excel)."""
    OUT_DIR.mkdir(exist_ok=True)
    stamp = biz_day or datetime.now().strftime("%Y-%m-%d")
    display = display_dt[DISPLAY_COLS].copy()
    for col in DATETIME_COLS:
        if col in display.columns:
            display[col] = parse_dt_series(display[col])

    csv_path = output_path_or_fallback(OUT_DIR / f"unified_master_{stamp}.csv")
    pq_path = output_path_or_fallback(OUT_DIR / f"unified_master_{stamp}.parquet")
    xlsx_path = output_path_or_fallback(OUT_DIR / f"unified_master_{stamp}.xlsx")

    print("  write Parquet", pq_path.name, flush=True)
    display.to_parquet(pq_path, index=False)

    if not full_excel:
        return xlsx_path, csv_path, pq_path

    print("  write dated CSV", csv_path.name, flush=True)
    display_str = display.copy()
    for col in DATETIME_COLS:
        if col in display_str.columns:
            display_str[col] = fmt_dt(display[col])
    display_str.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print("  write Excel", xlsx_path.name, flush=True)
    with pd.ExcelWriter(
        xlsx_path,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd hh:mm:ss",
    ) as writer:
        display.to_excel(writer, sheet_name="主明细", index=False)
        ws = writer.sheets["主明细"]
        ws.freeze_panes(1, 0)
        widths = {
            "A": 22, "B": 52, "C": 16, "D": 22, "E": 28, "F": 16,
            "G": 22, "H": 48, "I": 10, "J": 22, "K": 16, "L": 22,
            "M": 22, "N": 14, "O": 16, "P": 12, "Q": 12, "R": 14,
        }
        for col, w in widths.items():
            ws.set_column(f"{col}:{col}", w)
    return xlsx_path, csv_path, pq_path


def write_inter_hub_detail(
    master: pd.DataFrame, biz_day: str, hub_summary: pd.DataFrame | None = None,
) -> Path | None:
    stamp = biz_day or datetime.now().strftime("%Y-%m-%d")
    tagged = master[master["跨Hub转运"].astype("string").isin(["LAX-DFW", "EWR-ORD"])].copy()
    path = output_path_or_fallback(OUT_DIR / f"inter_hub_transfer_detail_{stamp}.xlsx")
    extra_dt = [
        "ags_handover_time", "ags_ascan_time", "actual_receiving_time",
        "recalculated_start_time", "recalculated_ascan_due_date",
    ]
    for col in INTER_HUB_RECALC_COLS:
        if col not in tagged.columns:
            tagged[col] = pd.NA
    extra_ags = (
        "ags_mawb_hub", "ags_tracking_origin_hub", "ags_destination_hub",
        "ags_destination_facility", "ags_handover_time", "ags_ascan_time",
    )
    cols = list(DISPLAY_COLS) + list(extra_ags) + list(INTER_HUB_RECALC_COLS)
    out = tagged.reindex(columns=cols)
    for col in DATETIME_COLS + extra_dt:
        if col in out.columns:
            out[col] = parse_dt_series(out[col])
    if hub_summary is None or hub_summary.empty:
        if len(tagged) and "ags_destination_hubs" in tagged.columns:
            hub_summary = tagged.groupby(
                ["跨Hub转运", "ags_mawb_hub", "MAWB"], dropna=False, as_index=False,
            ).agg(
                ags_destination_hubs=("ags_destination_hubs", "first"),
                tno_count=("TNO", "nunique"),
            )
        else:
            hub_summary = pd.DataFrame(columns=[
                "ags_mawb_hub", "mawb", "ags_destination_hubs", "跨Hub转运", "tno_count",
            ])
    summary_csv = output_path_or_fallback(OUT_DIR / f"inter_hub_mawb_summary_{stamp}.csv")
    hub_summary.to_csv(summary_csv, index=False, encoding="utf-8-sig")
    print(f"  write inter-hub detail {path.name} rows={len(out):,}", flush=True)
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        out.to_excel(writer, sheet_name="跨Hub转运明细", index=False)
        writer.sheets["跨Hub转运明细"].freeze_panes(1, 0)
        hub_summary.to_excel(writer, sheet_name="MAWB汇总", index=False)
        writer.sheets["MAWB汇总"].freeze_panes(1, 0)
    return path


def write_outputs(master: pd.DataFrame, audit: pd.DataFrame, mawb_conflicts: pd.DataFrame,
                  invalid: pd.DataFrame, files: dict, hub_summary: pd.DataFrame | None = None,
                  full_excel: bool = False):
    OUT_DIR.mkdir(exist_ok=True)
    master = add_ascan_due(master)
    if "CUSTOMER" not in master.columns and "Customer" in master.columns:
        master["CUSTOMER"] = master["Customer"]
    if "Customer" not in master.columns and "CUSTOMER" in master.columns:
        master["Customer"] = master["CUSTOMER"]
    if "cohort_date" not in master.columns:
        master["cohort_date"] = parse_dt_series(master["start_time"]).dt.strftime("%Y-%m-%d")

    if not audit.empty:
        audit = pd.concat([audit, mawb_conflicts], ignore_index=True) if len(mawb_conflicts) else audit
    elif len(mawb_conflicts):
        audit = mawb_conflicts.copy()
    else:
        audit = pd.DataFrame(columns=[
            "TNO", "field_name", "ags_raw", "system_raw", "platform_raw",
            "chosen_value", "chosen_reason", "needs_manual_review", "primary_source", "cohort_date",
        ])

    cohort_mode = master["cohort_date"].dropna().mode()
    biz_day = str(cohort_mode.iloc[0]) if len(cohort_mode) else None
    off = master[master["cohort_date"].notna() & (master["cohort_date"] != biz_day)].copy() if biz_day else master.iloc[0:0]

    display = master[DISPLAY_COLS].copy()
    write_display_files(display, biz_day, full_excel=full_excel)
    write_inter_hub_detail(master, biz_day, hub_summary)

    wide_keep = [c for c in FULL_COLS if c in master.columns]
    wide_pq = master[wide_keep].copy()
    for col in DATETIME_COLS:
        if col in wide_pq.columns:
            wide_pq[col] = parse_dt_series(wide_pq[col])
    wide_pq.to_parquet(output_path_or_fallback(OUT_DIR / "unified_master.parquet"), index=False)

    if full_excel:
        wide = wide_pq.copy()
        for col in DATETIME_COLS:
            if col in wide.columns:
                wide[col] = fmt_dt(wide[col])
        wide.to_csv(output_path_or_fallback(OUT_DIR / "unified_master.csv"), index=False, encoding="utf-8-sig")

    # keep audit / off-day / invalid as supporting files
    audit_csv = OUT_DIR / "conflict_audit.csv"
    off_csv = OUT_DIR / "off_day_records.csv"
    inv_csv = OUT_DIR / "invalid_no_tn.csv"
    audit.to_csv(audit_csv, index=False, encoding="utf-8-sig")
    if full_excel:
        off_disp = off[DISPLAY_COLS].copy() if len(off) else pd.DataFrame(columns=DISPLAY_COLS)
        for col in DATETIME_COLS:
            if col in off_disp.columns:
                off_disp[col] = fmt_dt(off_disp[col])
        off_disp.to_csv(off_csv, index=False, encoding="utf-8-sig")
        off_disp.to_csv(OUT_DIR / "out_of_cohort_records.csv", index=False, encoding="utf-8-sig")
    if invalid is None or invalid.empty:
        pd.DataFrame(columns=["TNO", "note"]).to_csv(inv_csv, index=False, encoding="utf-8-sig")
    else:
        invalid.to_csv(inv_csv, index=False, encoding="utf-8-sig")

    if not full_excel:
        return biz_day, master, audit, off

    print("  write Excel sample", flush=True)
    rng = np.random.default_rng(24)
    display_str = display.copy()
    for col in DATETIME_COLS:
        if col in display_str.columns:
            display_str[col] = fmt_dt(display_str[col])

    def sample_mask(mask, n=40):
        sub = display_str.loc[mask]
        if len(sub) <= n:
            return sub
        pick = rng.choice(sub.index.to_numpy(), size=n, replace=False)
        return sub.loc[pick]

    has_a = master["has_ags"] == "Yes" if "has_ags" in master.columns else pd.Series(False, index=master.index)
    has_s = master["has_system"] == "Yes" if "has_system" in master.columns else pd.Series(False, index=master.index)
    has_p = master["has_platform"] == "Yes" if "has_platform" in master.columns else pd.Series(False, index=master.index)
    plat = master["platform_sheet"] if "platform_sheet" in master.columns else pd.Series(pd.NA, index=master.index)
    sheets = {
        "三方重叠_TikTokGS": sample_mask(has_a & has_s & has_p, 40),
        "System加Platform_GS": sample_mask(has_s & has_p & ~has_a, 40),
        "AGS加AliExpress": sample_mask(has_a & has_p & ~has_s & (plat == "AliExpress"), 40),
        "System仅Pickup": sample_mask(has_s & ~has_a & ~has_p, 40),
        "AGS独有": sample_mask(has_a & ~has_s & ~has_p, 40),
        "FBT": sample_mask(plat == "FBT", 40),
        "Winit": sample_mask(plat == "Winit", 40),
        "PlatformGS独有": sample_mask((plat == "TT-GS") & ~has_s, 40),
        "跨Hub_LAX_DFW": sample_mask(master["跨Hub转运"].astype("string") == "LAX-DFW", 40) if "跨Hub转运" in master.columns else display_str.iloc[0:0],
        "跨Hub_EWR_ORD": sample_mask(master["跨Hub转运"].astype("string") == "EWR-ORD", 40) if "跨Hub转运" in master.columns else display_str.iloc[0:0],
    }
    sample_xlsx = output_path_or_fallback(OUT_DIR / "unified_master_sample.xlsx")
    with pd.ExcelWriter(sample_xlsx, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        pd.DataFrame([
            {"item": "business_day_mode", "value": biz_day},
            {"item": "master_rows", "value": len(display)},
            {"item": "unique_TNO", "value": display["TNO"].nunique()},
            {"item": "fbt_12h_rows", "value": int((master["sla_hours"] == 12).sum())},
            {"item": "procarrier_vinted_48h_rows", "value": int((master["sla_hours"] == 48).sum())},
            {"item": "other_24h_rows", "value": int((master["sla_hours"] == 24).sum())},
            {"item": "winit_calendar_d1_rows", "value": int((master["sla_type"].astype("string") == SLA_TYPE_CALENDAR_D1).sum()) if "sla_type" in master.columns else 0},
            {"item": "winit_calendar_d0_rows", "value": int((master["sla_type"].astype("string") == SLA_TYPE_CALENDAR_D0).sum()) if "sla_type" in master.columns else 0},
            {"item": "status_on_time", "value": int((master["status"] == "On Time").sum())},
            {"item": "status_late", "value": int((master["status"] == "Late").sum())},
            {"item": "status_no_ascan", "value": int((master["status"] == "No A-Scan").sum())},
            {"item": "inter_hub_lax_dfw", "value": int((master["跨Hub转运"].astype("string") == "LAX-DFW").sum()) if "跨Hub转运" in master.columns else 0},
            {"item": "inter_hub_ewr_ord", "value": int((master["跨Hub转运"].astype("string") == "EWR-ORD").sum()) if "跨Hub转运" in master.columns else 0},
            {"item": "ascan_rate_included", "value": int((~master["跨Hub转运"].astype("string").isin(["LAX-DFW", "EWR-ORD"])).sum()) if "跨Hub转运" in master.columns else len(master)},
        ]).to_excel(writer, sheet_name="摘要", index=False)
        pd.DataFrame({"column": DISPLAY_COLS}).to_excel(writer, sheet_name="字段顺序", index=False)
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)

    return biz_day, master, audit, off


def quality_report(master, audit, off, files, biz_day, ags_n, sys_n, plt_n):
    def nmask(m):
        return int(m.sum()) if hasattr(m, "sum") else int(m)

    has_a = master["has_ags"] == "Yes"
    has_s = master["has_system"] == "Yes"
    has_p = master["has_platform"] == "Yes"
    lines = []
    lines.append("# 统一主明细样例：数据质量报告")
    lines.append("")
    lines.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("原始文件只读，未修改、未重命名、未移动。")
    lines.append("")
    lines.append("> 本报告中的全部比率均为**辅助统计**，用于核对合并结果，不是正式 KPI 接口。")
    lines.append("> 正式 KPI 需在闭环程序中基于 `output/unified_master.csv` 宽表另行计算。")
    lines.append("")
    lines.append("## 输入识别")
    lines.append("")
    for kind in ("AGS", "System", "Platform"):
        path = files.get(kind)
        label = f"`{path.name}`" if path is not None else "**本批次缺席**"
        lines.append(f"- {kind}：{label}")
    lines.append("")
    lines.append("## 行数与去重")
    lines.append("")
    lines.append("| 项目 | 数量 |")
    lines.append("|---|---|")
    lines.append(f"| AGS 标准化行（按 TNO 去重后） | {ags_n:,} |")
    lines.append(f"| System 标准化行 | {sys_n:,} |")
    lines.append(f"| Platform 标准化行 | {plt_n:,} |")
    lines.append(f"| 合并后主表明细 | {len(master):,} |")
    lines.append(f"| 主表唯一 TNO | {master['TNO'].nunique():,} |")
    lines.append(f"| TNO 重复行 | {len(master) - master['TNO'].nunique():,} |")
    lines.append(f"| cohort_date 众数（业务日） | {biz_day} |")
    lines.append(f"| 非业务日记录（仍保留在主表） | {len(off):,} |")
    lines.append("")
    lines.append("## 来源重叠（合并后）")
    lines.append("")
    lines.append("| 组合 | 单量 |")
    lines.append("|---|---|")
    lines.append(f"| 仅 AGS | {nmask(has_a & ~has_s & ~has_p):,} |")
    lines.append(f"| 仅 System | {nmask(has_s & ~has_a & ~has_p):,} |")
    lines.append(f"| 仅 Platform | {nmask(has_p & ~has_a & ~has_s):,} |")
    lines.append(f"| AGS ∩ System（无 Platform） | {nmask(has_a & has_s & ~has_p):,} |")
    lines.append(f"| AGS ∩ Platform（无 System） | {nmask(has_a & has_p & ~has_s):,} |")
    lines.append(f"| System ∩ Platform（无 AGS） | {nmask(has_s & has_p & ~has_a):,} |")
    lines.append(f"| 三方 | {nmask(has_a & has_s & has_p):,} |")
    lines.append("")
    lines.append("## start_time_type")
    lines.append("")
    vc = master["start_time_type"].fillna("(空)").value_counts()
    lines.append("| start_time_type | 单量 |")
    lines.append("|---|---|")
    for k, v in vc.items():
        lines.append(f"| {k} | {int(v):,} |")
    lines.append("")
    lines.append("## CUSTOMER / Hub / primary_source")
    lines.append("")
    lines.append("### primary_source")
    lines.append("")
    lines.append("| primary_source | 单量 |")
    lines.append("|---|---|")
    for k, v in master["primary_source"].fillna("(空)").value_counts().items():
        lines.append(f"| {k} | {int(v):,} |")
    lines.append("")
    lines.append("### CUSTOMER（Top 15）")
    lines.append("")
    lines.append("| CUSTOMER | 单量 |")
    lines.append("|---|---|")
    cust_col = "CUSTOMER" if "CUSTOMER" in master.columns else "Customer"
    for k, v in master[cust_col].fillna("(空)").value_counts().head(15).items():
        lines.append(f"| {k} | {int(v):,} |")
    lines.append("")
    lines.append("### Hub（Top 15）")
    lines.append("")
    lines.append("| Hub | 单量 |")
    lines.append("|---|---|")
    for k, v in master["Hub"].fillna("(空)").value_counts().head(15).items():
        lines.append(f"| {k} | {int(v):,} |")
    lines.append("")
    lines.append("## 填充率")
    lines.append("")
    lines.append("| 字段 | 非空率 |")
    lines.append("|---|---|")
    for c in DISPLAY_COLS:
        if c in master.columns:
            rate = float(master[c].notna().mean())
            lines.append(f"| {c} | {rate:.2%} |")
    lines.append("")
    lines.append("## Ascan Due Date 与 status")
    lines.append("")
    if "sla_hours" in master.columns or "sla_type" in master.columns:
        lines.append("| SLA | 单量 |")
        lines.append("|---|---|")
        if "sla_type" in master.columns:
            for k, v in master["sla_type"].fillna("(空)").value_counts().items():
                lines.append(f"| sla_type={k} | {int(v):,} |")
        if "sla_hours" in master.columns:
            for k, v in master["sla_hours"].value_counts(dropna=True).items():
                lines.append(f"| +{int(k)} 小时 | {int(v):,} |")
        lines.append("")
    if "status" in master.columns:
        lines.append("| status | 单量 |")
        lines.append("|---|---|")
        for k, v in master["status"].fillna("(无法判断)").value_counts().items():
            lines.append(f"| {k} | {int(v):,} |")
        lines.append("")
    if "跨Hub转运" in master.columns:
        tagged = master["跨Hub转运"].astype("string").isin(["LAX-DFW", "EWR-ORD"])
        lax = master["跨Hub转运"].astype("string") == "LAX-DFW"
        ewr = master["跨Hub转运"].astype("string") == "EWR-ORD"
        lines.append("## 跨Hub转运与辅助 A-Scan Rate 范围")
        lines.append("")
        lines.append("跨 Hub 仅用 AGS 的 `MAWB Details.Hub` 与 `Tracking Details.destination_hub` 识别，再按 TNO 回写。不使用主表最终 `Hub`。")
        lines.append("")
        lines.append("| 项目 | 单量 |")
        lines.append("|---|---|")
        lines.append(f"| unified_master 总包裹量 | {len(master):,} |")
        lines.append(f"| LAX-DFW | {int(lax.sum()):,} |")
        lines.append(f"| EWR-ORD | {int(ewr.sum()):,} |")
        lines.append(f"| 因跨 Hub 暂不纳入辅助 Rate | {int(tagged.sum()):,} |")
        lines.append(f"| 纳入辅助 A-Scan Rate | {int((~tagged).sum()):,} |")
        lines.append("")
        included = master.loc[~tagged]
        if "status" in included.columns and len(included):
            lines.append("下面的辅助 Rate 仅统计 `跨Hub转运` 为空的包裹。")
            lines.append("")
            lines.append("| 纳入辅助 Rate 的 status | 单量 |")
            lines.append("|---|---|")
            for k, v in included["status"].fillna("(无法判断)").value_counts().items():
                lines.append(f"| {k} | {int(v):,} |")
            ot = int((included["status"] == "On Time").sum())
            late = int((included["status"] == "Late").sum())
            denom = ot + late
            rate = (ot / denom) if denom else None
            rate_txt = f"{rate:.2%}" if rate is not None else "N/A"
            lines.append("")
            lines.append(
                f"辅助 A-Scan On-Time Rate（On Time / (On Time + Late)，不含跨 Hub）= **{rate_txt}**。"
                "该数值仅供本次合并核对，**不是正式 KPI**，闭环程序不得直接引用。"
            )
            lines.append("")
        if "ags_mawb_hub" in master.columns:
            n_hub_diff = int(
                (master["ags_mawb_hub"].notna() & master["Hub"].notna()
                 & (master["ags_mawb_hub"].astype("string").str.upper() != master["Hub"].astype("string").str.upper())).sum()
            )
            lines.append(f"主表 Hub 与 ags_mawb_hub 不一致的行（预期可存在，因主表 Hub 走 Platform 优先）：{n_hub_diff:,}")
            lines.append("")
        tagged_rows = master.loc[tagged]
        if len(tagged_rows) and "ags_mawb_hub" in tagged_rows.columns:
            lines.append("### 转运 MAWB 清单")
            lines.append("")
            lines.append("| 跨Hub转运 | ags_mawb_hub | mawb | ags_destination_hubs | TNO |")
            lines.append("|---|---|---|---|---|")
            dest_col = "ags_destination_hubs" if "ags_destination_hubs" in tagged_rows.columns else "ags_destination_hub"
            mawb_tbl = tagged_rows.groupby(
                ["跨Hub转运", "ags_mawb_hub", "MAWB"], dropna=False, as_index=False,
            ).agg(
                ags_destination_hubs=(dest_col, "first"),
                tno_count=("TNO", "nunique"),
            ).sort_values(["跨Hub转运", "MAWB"])
            for _, r in mawb_tbl.iterrows():
                lines.append(
                    f"| {r['跨Hub转运']} | {r['ags_mawb_hub']} | {r['MAWB']} | {r['ags_destination_hubs']} | {int(r['tno_count']):,} |"
                )
            lines.append("")
            lax_mawb = int((mawb_tbl["跨Hub转运"] == "LAX-DFW").sum())
            ewr_mawb = int((mawb_tbl["跨Hub转运"] == "EWR-ORD").sum())
            lines.append(f"LAX-DFW MAWB 数 = {lax_mawb}；EWR-ORD MAWB 数 = {ewr_mawb}。")
            lines.append("")
    lines.append("## 冲突审计摘要")
    lines.append("")
    if audit.empty:
        lines.append("无冲突记录。")
    else:
        lines.append("| 字段 | 冲突条数 | 需人工 |")
        lines.append("|---|---|---|")
        g = audit.groupby("field_name")
        for field, sub in g:
            manual = int((sub["needs_manual_review"] == "Yes").sum())
            lines.append(f"| {field} | {len(sub):,} | {manual:,} |")
        lines.append("")
        lines.append("### 采用原因 Top")
        lines.append("")
        lines.append("| chosen_reason | 条数 |")
        lines.append("|---|---|")
        for k, v in audit["chosen_reason"].value_counts().head(12).items():
            lines.append(f"| {k} | {int(v):,} |")
    lines.append("")
    lines.append("## 映射核对要点（脱敏样例）")
    lines.append("")

    def show(df, title, cols=None):
        cols = cols or ["TNO", "Source", "业务名称", "CUSTOMER", "customer_raw", "start_time_type", "status"]
        lines.append(f"### {title}")
        lines.append("")
        if df.empty:
            lines.append("（无）")
            lines.append("")
            return
        show_df = df[cols].head(5).copy()
        show_df["TNO"] = show_df["TNO"].map(mask_tn)
        lines.append("| " + " | ".join(cols) + " |")
        lines.append("|" + "|".join(["---"] * len(cols)) + "|")
        for _, r in show_df.iterrows():
            vals = [str(r[c]) if pd.notna(r[c]) else "" for c in cols]
            lines.append("| " + " | ".join(vals) + " |")
        lines.append("")

    show_cols = ["TNO", "Source", "primary_source", "业务名称", "CUSTOMER", "customer_raw", "status"]
    show(master[has_a & has_s & has_p], "三方重叠（CUSTOMER/业务名称应为 TT GS，字段优先 Platform）", show_cols)
    show(master[has_a & has_p & ~has_s & (master["platform_sheet"] == "AliExpress")], "AGS ∩ AliExpress（CUSTOMER 应为 AliExpress）", show_cols)
    show(master[has_s & ~has_a & ~has_p], "System 独有 Pickup")
    show(master[master["platform_sheet"] == "FBT"], "FBT（CUSTOMER/业务名称应为 TT FBT，Hub=MIA 时 sla_hours=12，其他 Hub sla_hours=24，start_time_type Manifest Time）")
    show(master[master["platform_sheet"] == "Winit"], "Winit（CUSTOMER/业务名称应为 Winit，start_time_type Manifest Time，SLA 为日期制 D1 / LAX自提 D0）")

    lines.append("## 输出文件")
    lines.append("")
    lines.append("- `output/unified_master_YYYY-MM-DD.xlsx`：18 列主明细（含 跨Hub转运）")
    lines.append("- `output/inter_hub_transfer_detail_YYYY-MM-DD.xlsx`：跨 Hub 转运明细")
    lines.append("- `output/unified_master_YYYY-MM-DD.csv` / `.parquet`：同结构，供后续汇总")
    lines.append("- `output/unified_master_sample.xlsx`：按场景抽样")
    lines.append("- `output/conflict_audit.csv`：字段冲突审计")
    lines.append("- `output/off_day_records.csv`：交接时间日期 ≠ 业务日众数")
    lines.append("- `output/invalid_no_tn.csv`：空单号")
    lines.append("")
    lines.append("## 仍建议人工看一眼的点")
    lines.append("")
    lines.append("1. FBT/Winit 的 8/25 记录在主表中，同时在 off_day_records。")
    lines.append("2. Hub：System `first_ascan_hub` 优先于 Platform / AGS，不会被 Origin Hub 覆盖；没有 `first_ascan_hub` 时仍按 Platform → System `origin_poe` → AGS。")
    lines.append("3. AGS 的 TikTok 变体与 Cainiao 全称分别标准化为 `TT GS` / `AliExpress`；`customer_raw` 保留原值。")

    path = OUT_DIR / "data_quality_report.md"
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def normalize_master_aliases(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "start_time" not in df.columns and "交接时间" in df.columns:
        df["start_time"] = df["交接时间"]
    if "first_ascan_time" not in df.columns and "Ascan时间" in df.columns:
        df["first_ascan_time"] = df["Ascan时间"]
    if "start_time_type" not in df.columns and "交接时间类型" in df.columns:
        df["start_time_type"] = df["交接时间类型"]
    if "CUSTOMER" not in df.columns and "Customer" in df.columns:
        df["CUSTOMER"] = df["Customer"]
    if "Customer" not in df.columns and "CUSTOMER" in df.columns:
        df["Customer"] = df["CUSTOMER"]
    if "start_time_type" in df.columns:
        df["start_time_type"] = display_start_type(df["start_time_type"])
    if "交接时间类型" in df.columns:
        df["交接时间类型"] = display_start_type(df["交接时间类型"])
    if "跨Hub转运" not in df.columns:
        df["跨Hub转运"] = pd.NA
    return df


def export_from_merged_csv():
    """Rebuild dated 17-column outputs from the last merged CSV without rereading source Excel."""
    src = OUT_DIR / "unified_master.csv"
    if not src.exists():
        dated = sorted(OUT_DIR.glob("unified_master_*.csv"))
        if not dated:
            raise SystemExit("No merged CSV found in output/")
        src = dated[-1]
    print("load", src.name, flush=True)
    df = pd.read_csv(src, dtype={"TNO": "string", "MAWB": "string"})
    df = normalize_master_aliases(df)
    picked = df["CUSTOMER"] if "CUSTOMER" in df.columns else df["Customer"]
    has_ags = df["has_ags"].astype("string").eq("Yes") if "has_ags" in df.columns else pd.Series(False, index=df.index)
    ags_cust = df["customer_raw"] if "customer_raw" in df.columns else picked
    ags_cust = ags_cust.where(has_ags, pd.NA)
    _raw, std, ags_tt, ags_cainiao = standardize_ags_customer(picked, ags_cust)
    df["CUSTOMER"] = std
    df["Customer"] = std
    df["customer_standardized"] = std
    print(f"  AGS TikTok→TT GS={int(ags_tt.sum()):,} Cainiao→AliExpress={int(ags_cainiao.sum()):,}", flush=True)
    df = add_ascan_due(df)
    start = parse_dt_series(df["start_time"])
    df["cohort_date"] = start.dt.strftime("%Y-%m-%d")
    biz_day = str(df["cohort_date"].dropna().mode().iloc[0])
    write_display_files(df[DISPLAY_COLS], biz_day)

    off = df[df["cohort_date"].notna() & (df["cohort_date"] != biz_day)].copy()
    off_disp = off[DISPLAY_COLS].copy() if len(off) else pd.DataFrame(columns=DISPLAY_COLS)
    for col in DATETIME_COLS:
        if col in off_disp.columns:
            off_disp[col] = fmt_dt(off_disp[col])
    off_disp.to_csv(OUT_DIR / "off_day_records.csv", index=False, encoding="utf-8-sig")
    off_disp.to_csv(OUT_DIR / "out_of_cohort_records.csv", index=False, encoding="utf-8-sig")
    print(f"  out_of_cohort_records {len(off_disp):,}", flush=True)

    print("DONE export", biz_day, "rows", len(df), "unique", df["TNO"].nunique(), flush=True)
    print(df["status"].fillna("(na)").value_counts().to_dict(), flush=True)
    print("  columns", list(df[DISPLAY_COLS].columns), flush=True)
    print("  start_time_type", df["start_time_type"].fillna("(空)").value_counts().to_dict(), flush=True)
    winit = df["CUSTOMER"].astype("string") == "Winit"
    fbt = df["CUSTOMER"].astype("string") == "TT FBT"
    winit_type = df.loc[winit, "sla_type"].value_counts().to_dict() if "sla_type" in df.columns else {}
    print(f"  Winit={int(winit.sum()):,} type={df.loc[winit, 'start_time_type'].value_counts().to_dict()} sla_type={winit_type}", flush=True)
    print(f"  FBT={int(fbt.sum()):,} sla={df.loc[fbt, 'sla_hours'].value_counts().to_dict()}", flush=True)
    sla48 = is_48h_sla_row(df)
    print(
        f"  ProCarrier/Vinted 48h={int(sla48.sum()):,} sla={df.loc[sla48, 'sla_hours'].value_counts().to_dict()}",
        flush=True,
    )
    print(f"  Pending left={int((df['status'].astype('string') == 'Pending').sum()):,}", flush=True)
    print(f"  date_manifested left={int((df['start_time_type'].astype('string') == 'date_manifested').sum()):,}", flush=True)


def verify_tno_coverage(master: pd.DataFrame, ags: pd.DataFrame, sys_df: pd.DataFrame, plt: pd.DataFrame) -> None:
    """Every valid input TNO must land in the master exactly once, and nothing extra."""
    print("VERIFY TNO coverage", flush=True)
    sets = {
        "AGS": set(ags["TNO"].dropna().astype(str)),
        "System": set(sys_df["TNO"].dropna().astype(str)),
        "Platform": set(plt["TNO"].dropna().astype(str)),
    }
    for name, s in sets.items():
        print(f"  {name} unique TNO = {len(s):,}", flush=True)
    union = set().union(*sets.values())
    got = set(master["TNO"].dropna().astype(str))
    missing = union - got
    extra = got - union
    print(f"  input union = {len(union):,}", flush=True)
    print(f"  master rows = {len(master):,} unique = {master['TNO'].nunique():,} "
          f"dup = {len(master) - master['TNO'].nunique():,}", flush=True)
    print(f"  missing from master = {len(missing):,}", flush=True)
    print(f"  in master but not in any input = {len(extra):,}", flush=True)
    if missing:
        print(f"    missing sample: {sorted(missing)[:10]}", flush=True)
    if extra:
        print(f"    extra sample: {sorted(extra)[:10]}", flush=True)
    if "platform_sheet" in master.columns:
        vc = master["platform_sheet"].dropna().value_counts().to_dict()
        print(f"  master rows per platform_sheet: {vc}", flush=True)


def verify_platform_priority(master: pd.DataFrame) -> None:
    """Print checks for Platform > System > AGS field priority."""
    print("VERIFY platform field priority", flush=True)
    n = len(master)
    uniq = master["TNO"].nunique()
    print(f"  rows={n:,} unique_TNO={uniq:,} dup={n - uniq:,}", flush=True)

    has_p = master["has_platform"] == "Yes"
    gs = has_p & (master["platform_sheet"].astype("string") == "TT-GS")
    ali = has_p & (master["platform_sheet"].astype("string") == "AliExpress")
    src_has_plt = master["Source"].astype("string").str.contains("Platform", na=False)

    gs_cust = master.loc[gs, "CUSTOMER"].fillna("(空)").astype(str) if "CUSTOMER" in master.columns else master.loc[gs, "Customer"].fillna("(空)").astype(str)
    gs_biz = master.loc[gs, "业务名称"].fillna("(空)").astype(str)
    bad_gs_cust = int((gs_cust != "TT GS").sum())
    bad_gs_biz = int((gs_biz != "TT GS").sum())
    tiktok_inc_on_gs = int(gs_cust.str.contains("TikTok Inc|TIKTOK", case=False, regex=True).sum())
    print(f"  TT-GS rows={int(gs.sum()):,} Customer!=TT GS={bad_gs_cust:,} 业务名称!=TT GS={bad_gs_biz:,} TikTok/TIKTOK on GS={tiktok_inc_on_gs:,}", flush=True)

    ali_cust = master.loc[ali, "CUSTOMER"].fillna("(空)").astype(str) if "CUSTOMER" in master.columns else master.loc[ali, "Customer"].fillna("(空)").astype(str)
    ali_biz = master.loc[ali, "业务名称"].fillna("(空)").astype(str)
    print(f"  AliExpress rows={int(ali.sum()):,} Customer!=AliExpress={int((ali_cust != 'AliExpress').sum()):,} 业务名称!=AliExpress={int((ali_biz != 'AliExpress').sum()):,}", flush=True)

    src_plt_not_gs = src_has_plt & gs
    print(f"  Source contains Platform and TT-GS: {int(src_plt_not_gs.sum()):,}", flush=True)

    plt_rows = master.loc[has_p]
    for col in ("customer_source", "hub_source", "start_time_source", "ascan_source", "mawb_source", "primary_source"):
        if col not in master.columns:
            continue
        vc = plt_rows[col].fillna("(空)").value_counts()
        print(f"  {col} on Platform rows: {vc.to_dict()}", flush=True)

    fill_sys_ags = has_p & (
        master.get("customer_source", pd.Series(index=master.index)).isin(["System", "AGS"])
        | master.get("hub_source", pd.Series(index=master.index)).isin(["System", "AGS"])
        | master.get("mawb_source", pd.Series(index=master.index)).isin(["System", "AGS"])
        | master.get("start_time_source", pd.Series(index=master.index)).isin(["System", "AGS"])
        | master.get("ascan_source", pd.Series(index=master.index)).isin(["System", "AGS"])
    )
    print(f"  Platform rows with System/AGS fill: {int(fill_sys_ags.sum()):,}", flush=True)

    if "customer_raw" in master.columns:
        has_a = master["has_ags"] == "Yes"
        raw_tt = has_a & is_tiktok_customer_series(master["customer_raw"])
        raw_cainiao = has_a & is_cainiao_customer_series(master["customer_raw"])
        std = master["CUSTOMER"].astype("string") if "CUSTOMER" in master.columns else master["Customer"].astype("string")
        print(f"  AGS TikTok raw={int(raw_tt.sum()):,} display!=TT GS={int((raw_tt & (std != 'TT GS')).sum()):,}", flush=True)
        print(f"  AGS Cainiao raw={int(raw_cainiao.sum()):,} display!=AliExpress={int((raw_cainiao & (std != 'AliExpress')).sum()):,}", flush=True)
        print(f"  display CUSTOMER contains TikTok={int(std.str.contains(TIKTOK_RE, na=False).sum()):,}", flush=True)
        print(f"  display CUSTOMER contains CAINIAO={int(std.str.contains('CAINIAO', case=False, na=False).sum()):,}", flush=True)
    print("  DISPLAY_COLS", DISPLAY_COLS, flush=True)
    if "status" in master.columns:
        print("  status", master["status"].fillna("(空)").value_counts().to_dict(), flush=True)
        print(f"  Pending left={int((master['status'].astype('string') == 'Pending').sum()):,}", flush=True)
    if "sla_hours" in master.columns:
        print("  sla_hours", master["sla_hours"].value_counts(dropna=False).to_dict(), flush=True)
        sla48 = is_48h_sla_row(master)
        print(f"  ProCarrier/Vinted System-only 48h={int(sla48.sum()):,}", flush=True)
    if "sla_type" in master.columns:
        print("  sla_type", master["sla_type"].fillna("(空)").value_counts().to_dict(), flush=True)
    if "start_time_type" in master.columns:
        print("  start_time_type", master["start_time_type"].fillna("(空)").value_counts().to_dict(), flush=True)
    if "跨Hub转运" in master.columns:
        tag = master["跨Hub转运"].astype("string")
        print("  跨Hub转运", tag.fillna("(空)").value_counts().head(8).to_dict(), flush=True)
        unexpected = tag.dropna()
        unexpected = unexpected[~unexpected.isin(["LAX-DFW", "EWR-ORD"])]
        print(f"  unexpected 跨Hub转运 values={int(len(unexpected)):,}", flush=True)
        tagged = tag.isin(["LAX-DFW", "EWR-ORD"])
        non_ags = tagged & (master.get("has_ags", pd.Series("No", index=master.index)) != "Yes")
        print(f"  tagged non-AGS TNO={int(non_ags.sum()):,} (must be 0)", flush=True)
        print(f"  unique TNO={master['TNO'].nunique():,} rows={len(master):,}", flush=True)
        if "ags_mawb_hub" in master.columns:
            same = int(
                (master["ags_mawb_hub"].notna() & master["Hub"].notna()
                 & (master["Hub"].astype("string").str.upper() == master["ags_mawb_hub"].astype("string").str.upper())).sum()
            )
            diff = int(
                (master["ags_mawb_hub"].notna() & master["Hub"].notna()
                 & (master["Hub"].astype("string").str.upper() != master["ags_mawb_hub"].astype("string").str.upper())).sum()
            )
            print(f"  Hub equals ags_mawb_hub={same:,}; Hub differs from ags_mawb_hub={diff:,} (Hub not overwritten)", flush=True)
            tagged_hub_src = master.loc[tagged, "hub_source"].value_counts().to_dict() if "hub_source" in master.columns else {}
            print(f"  tagged hub_source {tagged_hub_src}", flush=True)


def main():
    if "--from-csv" in sys.argv:
        export_from_merged_csv()
        return
    full_excel = "--full-excel" in sys.argv
    print("discover files", flush=True)
    files = find_input_files()
    print(files, flush=True)
    if "AGS" in files:
        ags, inv_a, mawb_c = process_ags(files["AGS"])
        ags, hub_summary = tag_ags_inter_hub(ags)
    else:
        print("  AGS skipped (file missing this run)", flush=True)
        ags = empty_std()
        for col in AGS_EXTRA_COLS:
            if col not in ags.columns:
                ags[col] = pd.NA
        inv_a = pd.DataFrame()
        mawb_c = pd.DataFrame()
        hub_summary = pd.DataFrame()
    if "System" in files:
        sys_df, inv_s = process_system(files["System"])
    else:
        print("  System skipped (file missing this run)", flush=True)
        sys_df, inv_s = empty_std(), pd.DataFrame()
    if "Platform" in files:
        plt, inv_p = process_platform(files["Platform"])
    else:
        print("  Platform skipped (file missing this run)", flush=True)
        plt, inv_p = empty_std(), pd.DataFrame()
    invalid_frames = [x for x in (inv_a, inv_s, inv_p) if x is not None and not x.empty]
    invalid = pd.concat(invalid_frames, ignore_index=True) if invalid_frames else pd.DataFrame()
    master, audit = merge_sources(ags, sys_df, plt)
    master = attach_ags_inter_hub(master, ags)
    biz_day, master_out, audit_out, off = write_outputs(
        master, audit, mawb_c, invalid, files, hub_summary, full_excel=full_excel,
    )
    qpath = quality_report(master_out, audit_out, off, files, biz_day, len(ags), len(sys_df), len(plt))
    verify_tno_coverage(master_out, ags, sys_df, plt)
    verify_platform_priority(master_out)
    print("DONE", flush=True)
    print("master", len(master_out), "unique", master_out["TNO"].nunique(), flush=True)
    print("audit", len(audit_out), "off_day", len(off), flush=True)
    print("quality", qpath, flush=True)


if __name__ == "__main__":
    main()
