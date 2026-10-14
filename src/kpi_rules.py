# -*- coding: utf-8 -*-
"""KPI timing and status rules. SLA reuses Data Combination so there is one definition."""
from __future__ import annotations

import re

import pandas as pd

from common import clean_text, norm_key
from dc_bridge import merge_module

STATUS_ON_TIME = "On Time"
STATUS_LATE = "Late"
STATUS_LATE_PICKUP = "Late(PickUp)"
STATUS_LATE_HANDOVER = "Late(Handover)"
STATUS_LATE_MANIFEST = "Late(Manifest)"
STATUS_LATE_OTHERS = "Late(Others)"


def sla_hours_for(df: pd.DataFrame) -> pd.Series:
    """Hour-based SLA only. Winit calendar rows are NA so they cannot be judged in hours."""
    module = merge_module()
    start = pd.to_datetime(df["start_time"], errors="coerce") if "start_time" in df.columns else pd.Series(
        pd.NaT, index=df.index
    )
    return module.compute_sla_fields(df, start)["sla_hours"]


def resolve_sla_hours(df: pd.DataFrame) -> pd.Series:
    """Recompute rather than trust a leftover 24h on Winit rows."""
    return sla_hours_for(df)


def timezone_exceptions(series: pd.Series) -> pd.Series:
    """Flag tz-aware or tz-suffixed values instead of converting them."""
    if pd.api.types.is_datetime64tz_dtype(series):
        return pd.Series(True, index=series.index)
    text = series.astype("string")
    return text.str.contains(r"(?:[+-]\d{2}:?\d{2}|Z|UTC|GMT)\s*$", na=False, regex=True)


def compute_kpi_times(df: pd.DataFrame, kpi_start: pd.Series) -> pd.DataFrame:
    """kpi_date is the A-Scan due date (应履约日期), not the scan or run date.

    Hour customers: due = kpi_start + SLA hours.
    Winit: due from Manifested Time (start_time), calendar D1 / LAX自提 D0.
    Receiving time does not change Winit due dates.
    """
    module = merge_module()
    clock = pd.to_datetime(kpi_start, errors="coerce")
    winit = module.is_winit_row(df)
    if "start_time" in df.columns:
        manifest = pd.to_datetime(df["start_time"], errors="coerce")
        clock = clock.mask(winit, manifest)
    spec = module.compute_sla_fields(df, clock)
    out = pd.DataFrame(index=df.index)
    # Float64 so Pool Excel (float64) accepts 12/24/48 and Winit NA together.
    out["sla_hours"] = pd.to_numeric(spec["sla_hours"], errors="coerce").astype("float64")
    out["sla_type"] = spec["sla_type"]
    out["kpi_start_time"] = clock
    out["kpi_due_date"] = spec["due_display"]
    out["kpi_date"] = spec["kpi_date"]
    out["kpi_deadline"] = spec["deadline"]
    return out


def classify_base_status(ascan: pd.Series, time_diff: pd.Series, sla_hours: pd.Series,
                         sla_type: pd.Series | None = None,
                         kpi_due_date: pd.Series | None = None,
                         kpi_deadline: pd.Series | None = None) -> pd.Series:
    """Empty A-scan -> No A-scan. Hour SLA uses time_diff vs sla_hours.
    Calendar SLA uses ascan vs next-day 00:00:00. time_diff is never the Winit test.
    """
    module = merge_module()
    ascan = pd.to_datetime(ascan, errors="coerce")
    status = pd.Series(pd.NA, index=ascan.index, dtype="object")
    has_ascan = ascan.notna()

    if sla_type is None:
        sla_type = pd.Series(module.SLA_TYPE_HOURS, index=ascan.index, dtype="object")
    sla_type = sla_type.astype("string")
    cal = sla_type.isin(list(module.CALENDAR_SLA_TYPES))

    if kpi_deadline is None and kpi_due_date is not None:
        due = pd.to_datetime(kpi_due_date, errors="coerce")
        kpi_deadline = due.dt.normalize() + pd.Timedelta(days=1)
        kpi_deadline = kpi_deadline.mask(~cal, due)

    if kpi_deadline is not None:
        on_time = module.ascan_on_time(ascan, kpi_deadline, cal)
        late = has_ascan & pd.to_datetime(kpi_deadline, errors="coerce").notna() & ~on_time
        status = status.mask(on_time, STATUS_ON_TIME)
        status = status.mask(late, STATUS_LATE)
    else:
        diff = pd.to_numeric(time_diff, errors="coerce")
        sla = pd.to_numeric(sla_hours, errors="coerce")
        comparable = has_ascan & diff.notna() & sla.notna() & ~cal
        status = status.mask(comparable & (diff <= sla), STATUS_ON_TIME)
        status = status.mask(comparable & (diff > sla), STATUS_LATE)

    status = status.mask(~has_ascan, "__NO_ASCAN__")
    return status


def subclass_no_ascan(status: pd.Series, queried: pd.Series, event_code: pd.Series, cfg: dict) -> pd.Series:
    """After A-scan is still empty: empty code, 51005, and other codes are three classes.

    Not queried is never treated as an empty event code.
    """
    sub = cfg["no_ascan_subclass"]
    code_text = event_code.map(lambda v: clean_text(v))
    has_code = code_text.notna()
    hit_51005 = _event_code_hit(event_code, sub.get("event_codes_51005") or ["51005"])
    out = status.copy()
    no_ascan = status.astype("string") == "__NO_ASCAN__"
    out = out.mask(no_ascan & ~queried, sub["not_queried"])
    out = out.mask(no_ascan & queried & ~has_code, sub["no_track"])
    out = out.mask(no_ascan & queried & hit_51005, sub.get("status_51005", "51005"))
    out = out.mask(no_ascan & queried & has_code & ~hit_51005, sub["missing_ascan"])
    return out


def start_type_series(df: pd.DataFrame) -> pd.Series:
    if "start_time_type" in df.columns:
        return df["start_time_type"]
    if "交接时间类型" in df.columns:
        return df["交接时间类型"]
    return pd.Series(pd.NA, index=df.index)


def _start_type_key(series: pd.Series) -> pd.Series:
    return series.astype("string").str.lower().str.replace(r"[\s_\-]+", "", regex=True)


def is_pickup_time_type(series: pd.Series) -> pd.Series:
    """True when start_time_type is pickup-related (Pickup Time, Pickup, Pick Up, ...)."""
    return _start_type_key(series).str.contains("pickup", na=False)


def is_handover_time_type(series: pd.Series) -> pd.Series:
    """True when start_time_type is handover-related (Handover Time, Handover, ...)."""
    return _start_type_key(series).str.contains("handover", na=False)


def is_manifest_time_type(series: pd.Series) -> pd.Series:
    """True when start_time_type is manifest-related (Manifest Time, date_manifested, ...)."""
    return _start_type_key(series).str.contains("manifest", na=False)


def split_late(status: pd.Series, start_type: pd.Series, cfg: dict | None = None) -> pd.Series:
    """Bare Late / leftover Late(Others) -> PickUp / Handover / Manifest by start_time_type.

    Pickup wins if a value matches more than one token. Blank or unknown types stay
    Late(Others) so they are not silently folded into Handover or Manifest.
    Already-split PickUp / Handover / Manifest labels stay put.
    """
    split = (cfg or {}).get("late_split") or {}
    pickup_label = split.get("pickup", STATUS_LATE_PICKUP)
    handover_label = split.get("handover", STATUS_LATE_HANDOVER)
    manifest_label = split.get("manifest", STATUS_LATE_MANIFEST)
    others_label = split.get("others", STATUS_LATE_OTHERS)
    text = status.astype("string")
    need = text.eq(STATUS_LATE) | text.eq(others_label)
    pickup = is_pickup_time_type(start_type).reindex(status.index).fillna(False)
    handover = is_handover_time_type(start_type).reindex(status.index).fillna(False)
    manifest = is_manifest_time_type(start_type).reindex(status.index).fillna(False)
    out = status.copy()
    out = out.mask(need & pickup, pickup_label)
    out = out.mask(need & ~pickup & handover, handover_label)
    out = out.mask(need & ~pickup & ~handover & manifest, manifest_label)
    out = out.mask(need & ~pickup & ~handover & ~manifest, others_label)
    return out


def _norm_desc(value) -> str:
    text = clean_text(value)
    if text is None:
        return ""
    return re.sub(r"\s+", " ", text).strip().lower()


def _event_code_hit(series: pd.Series, values: list) -> pd.Series:
    codes = {norm_key(c) for c in values}
    code_norm = series.map(lambda v: norm_key(v) if clean_text(v) is not None else "")
    # Tolerate a numeric export turning 57407 into 57407.0
    code_norm = code_norm.str.replace(r"\.0+$", "", regex=True)
    return code_norm.isin(codes) & code_norm.ne("")


def _contains_any(series: pd.Series, values: list) -> pd.Series:
    text = series.map(_norm_desc)
    hit = pd.Series(False, index=series.index)
    for phrase in values:
        token = _norm_desc(phrase)
        if not token:
            continue
        hit |= text.str.contains(re.escape(token), na=False)
    return hit


def wrong_address_match(fields: dict[str, pd.Series], cfg: dict) -> tuple[pd.Series, pd.Series]:
    """Any configured rule on Status Check columns. Returns (hit mask, source labels)."""
    rules = cfg["wrong_address"]
    index = next(iter(fields.values())).index if fields else pd.RangeIndex(0)
    hit = pd.Series(False, index=index)
    sources: list[pd.Series] = []
    for rule in rules.get("rules") or []:
        field = rule.get("field")
        series = fields.get(field)
        if series is None:
            continue
        match_type = rule.get("match_type")
        values = rule.get("values") or []
        if match_type == "event_code":
            part = _event_code_hit(series, values)
        elif match_type == "not_empty":
            part = series.map(clean_text).notna()
        else:
            part = _contains_any(series, values)
        label = rule.get("source") or field or "rule"
        sources.append(part.map(lambda flag, name=label: name if flag else ""))
        hit = hit | part
    if sources:
        joined = pd.concat(sources, axis=1).apply(
            lambda row: "; ".join(part for part in row if part), axis=1
        )
    else:
        joined = pd.Series("", index=index)
    return hit, joined


def apply_wrong_address(status: pd.Series, flag: pd.Series, cfg: dict) -> tuple[pd.Series, pd.Series]:
    rules = cfg["wrong_address"]
    label = rules.get("label", "Wrong Address")
    allowed = set(rules.get("can_override", []))
    blocked = set(rules.get("never_override", []))
    current = status.astype("string")
    can = flag & current.isin(allowed) & ~current.isin(blocked)
    return status.mask(can, label), flag & ~can


def on_time_count(status: pd.Series) -> pd.Series:
    return (status.astype("string") == STATUS_ON_TIME).astype("int64")


ALIEXPRESS_NAMES = {"aliexpress"}
NO_AGS_EXCLUDE_HUBS = {"MIA"}
ORD_HUB = "ORD"
PROCARRIER_TOKEN = "procarrier"
ELIGIBLE_TRUE = "True"
ELIGIBLE_FALSE = "False"


def coerce_eligible(series: pd.Series) -> pd.Series:
    """Accept bool / Yes / True from parquet or a rebuilt report."""
    if series is None:
        return pd.Series(dtype="boolean")
    text = series.astype("string").str.strip().str.lower()
    return text.isin(["true", "yes", "1"])


def _name_is_aliexpress(df: pd.DataFrame, columns: tuple[str, ...]) -> pd.Series:
    hit = pd.Series(False, index=df.index)
    for col in columns:
        if col not in df.columns:
            continue
        part = df[col].astype("string").str.strip().str.lower().isin(ALIEXPRESS_NAMES)
        hit = hit | part.fillna(False)
    return hit


def is_aliexpress_name_or_customer(df: pd.DataFrame) -> pd.Series:
    """业务名称 or CUSTOMER/Customer equals AliExpress (case-insensitive)."""
    return _name_is_aliexpress(df, ("业务名称", "CUSTOMER", "Customer"))


def is_aliexpress_row(df: pd.DataFrame) -> pd.Series:
    return _name_is_aliexpress(
        df, ("CUSTOMER", "Customer", "业务名称", "platform_sheet", "customer_standardized")
    )


def source_has_ags(df: pd.DataFrame) -> pd.Series:
    """True only when the row actually came from AGS. Pool labels do not count."""
    if "has_ags" in df.columns:
        known = df["has_ags"].astype("string").str.strip().str.lower()
        flag = known.eq("yes")
        missing = known.isna() | known.eq("") | known.eq("<na>") | known.eq("nan")
        if not bool(missing.any()):
            return flag
    else:
        flag = pd.Series(False, index=df.index)
        missing = pd.Series(True, index=df.index)
    text = pd.Series("", index=df.index, dtype="object")
    for col in ("original_source", "Source"):
        if col in df.columns:
            part = df[col].astype("string").fillna("")
            text = text.where(text.astype("string").str.len() > 0, part)
    fallback = text.astype("string").str.contains(r"(?i)(?:^|[^a-z])ags(?:[^a-z]|$)", na=False)
    return flag.where(~missing, fallback)


def _hub_series(df: pd.DataFrame) -> pd.Series:
    if "Hub" not in df.columns:
        return pd.Series("", index=df.index, dtype="string")
    return df["Hub"].astype("string").str.strip().str.upper()


def _norm_customer_token(series: pd.Series) -> pd.Series:
    t = series.astype("string").str.strip().str.lower()
    return t.str.replace(r"[\s_\-]+", "", regex=True)


def is_procarrier_customer(df: pd.DataFrame) -> pd.Series:
    """CUSTOMER is ProCarrier. Pro Carrier T11 is a different name and stays eligible."""
    hit = pd.Series(False, index=df.index)
    for col in ("CUSTOMER", "Customer", "业务名称"):
        if col not in df.columns:
            continue
        # Pool-only rows may lack Customer; False | <NA> must not become <NA>.
        hit = hit | _norm_customer_token(df[col]).eq(PROCARRIER_TOKEN).fillna(False)
    return hit


def aliexpress_ord_ineligible(df: pd.DataFrame) -> pd.Series:
    """业务名称 or customer is AliExpress and Hub=ORD: keep the row, leave KPI."""
    if df.empty:
        return pd.Series(False, index=df.index)
    return is_aliexpress_name_or_customer(df) & _hub_series(df).eq(ORD_HUB)


def aliexpress_no_ags_mia(df: pd.DataFrame) -> pd.Series:
    """AliExpress with no AGS source at MIA: keep the row, leave KPI."""
    if df.empty:
        return pd.Series(False, index=df.index)
    return is_aliexpress_row(df) & ~source_has_ags(df) & _hub_series(df).isin(NO_AGS_EXCLUDE_HUBS)


def kpi_ineligible_mask(df: pd.DataFrame) -> pd.Series:
    return aliexpress_ord_ineligible(df) | aliexpress_no_ags_mia(df) | is_procarrier_customer(df)


def aliexpress_no_ags_mia_ord(df: pd.DataFrame) -> pd.Series:
    """Full ineligible mask. Name kept for existing call sites."""
    return kpi_ineligible_mask(df)


def apply_eligibility_exclusions(eligible: pd.Series, df: pd.DataFrame) -> pd.Series:
    """Only flip kpi_eligible. Does not drop rows or change SLA."""
    return coerce_eligible(eligible) & ~kpi_ineligible_mask(df).fillna(False)


REASON_PROCARRIER = "ProCarrier"
REASON_VINTED = "Vinted"
REASON_DFW_FBT = "DFW 的 TT FBT"
REASON_IAH_HANDOVER = "IAH 且 Handover"
REMOVAL_REASONS = (REASON_PROCARRIER, REASON_VINTED, REASON_DFW_FBT, REASON_IAH_HANDOVER)
REMOVAL_NOTES = {
    REASON_PROCARRIER: "CUSTOMER 或业务名称是 ProCarrier。Pro Carrier T11 保留。",
    REASON_VINTED: "CUSTOMER 或业务名称包含 Vinted。",
    REASON_DFW_FBT: "只删当前 Hub 仍是 DFW 的 TT FBT。IAH 的 TT FBT 保留，包括合并时从 DFW 改到 IAH 的。",
    REASON_IAH_HANDOVER: "Hub 为 IAH 且 start_time_type 含 Handover。IAH 的 TT FBT 若不是 Handover 则保留。",
}
REMOVAL_DETAIL_COLS = (
    "删除原因", "删除范围", "TNO", "MAWB", "CUSTOMER", "业务名称", "Hub",
    "start_time_type", "start_time", "first_ascan_time", "status", "Source",
    "进表来源", "kpi_eligible", "kpi_date", "sla_hours", "sla_type", "merge_note",
    "query_reason",
)


def is_vinted_customer(df: pd.DataFrame) -> pd.Series:
    """CUSTOMER or 业务名称 contains Vinted. Pro Carrier T11 is not Vinted."""
    hit = pd.Series(False, index=df.index)
    for col in ("CUSTOMER", "Customer", "业务名称"):
        if col not in df.columns:
            continue
        text = df[col].astype("string").str.strip().str.lower()
        hit = hit | text.str.contains("vinted", na=False).fillna(False)
    return hit


def is_dfw_ttfbt(df: pd.DataFrame) -> pd.Series:
    """TT FBT whose Hub is still DFW. IAH TT FBT stays, including DFW rewritten to IAH."""
    if df is None or df.empty:
        return pd.Series(False, index=df.index if df is not None else [])
    fbt = merge_module().is_fbt_row(df).fillna(False)
    return fbt & _hub_series(df).eq("DFW").fillna(False)


def removal_reason(df: pd.DataFrame) -> pd.Series:
    """One reason per row. ProCarrier, then Vinted, then DFW TT FBT, then IAH Handover."""
    reason = pd.Series(pd.NA, index=df.index, dtype="object")
    if df is None or df.empty:
        return reason
    dfw = is_dfw_ttfbt(df)
    iah = (
        _hub_series(df).eq("IAH").fillna(False)
        & is_handover_time_type(start_type_series(df)).fillna(False)
        & ~dfw
    )
    reason = reason.mask(iah, REASON_IAH_HANDOVER)
    reason = reason.mask(dfw, REASON_DFW_FBT)
    reason = reason.mask(is_vinted_customer(df), REASON_VINTED)
    reason = reason.mask(is_procarrier_customer(df), REASON_PROCARRIER)
    return reason


def split_removed_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Drop ProCarrier, Vinted, DFW TT FBT, and IAH Handover. Kept rows are unchanged."""
    if df is None:
        empty = pd.DataFrame(columns=["删除原因"])
        return empty, empty
    reason = removal_reason(df)
    hit = reason.notna()
    removed = df.loc[hit].copy()
    removed["删除原因"] = reason.loc[hit].astype("object")
    kept = df.loc[~hit].copy()
    return kept, removed


def removal_summary(removed: pd.DataFrame, before_rows: int, kept_rows: int,
                    kept_eligible: int) -> pd.DataFrame:
    rows = []
    for name in REMOVAL_REASONS:
        n = int((removed["删除原因"] == name).sum()) if len(removed) and "删除原因" in removed.columns else 0
        rows.append({"项目": name, "条数": n, "说明": REMOVAL_NOTES[name]})
    rows.append({"项目": "合计删除", "条数": int(len(removed)), "说明": "同一条只归一个原因"})
    rows.append({"项目": "删除前主表", "条数": int(before_rows), "说明": ""})
    rows.append({"项目": "删除后主表", "条数": int(kept_rows), "说明": ""})
    rows.append({
        "项目": "删除后计入 KPI",
        "条数": int(kept_eligible),
        "说明": "kpi_eligible=True，作为本次分母",
    })
    return pd.DataFrame(rows)


def removal_detail(removed: pd.DataFrame) -> pd.DataFrame:
    if removed is None or removed.empty:
        return pd.DataFrame(columns=list(REMOVAL_DETAIL_COLS))
    cols = [c for c in REMOVAL_DETAIL_COLS if c in removed.columns]
    return removed.loc[:, cols].copy()


def apply_hub_corrections(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Reuse Data Combination Hub rewrites so KPI / Pool see the same Hub."""
    if df is None or df.empty:
        return df, 0
    return merge_module().remap_fbt_dfw_hub(df)


def eligible_label(flag: pd.Series) -> pd.Series:
    return coerce_eligible(flag).map({True: ELIGIBLE_TRUE, False: ELIGIBLE_FALSE})


def time_diff_hours(ascan: pd.Series, kpi_start: pd.Series) -> pd.Series:
    a = pd.to_datetime(ascan, errors="coerce")
    s = pd.to_datetime(kpi_start, errors="coerce")
    return ((a - s) / pd.Timedelta(hours=1)).round(4)
