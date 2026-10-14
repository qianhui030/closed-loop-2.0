# -*- coding: utf-8 -*-
"""Step 1: build the queryable Unified Master and the status query list."""
from __future__ import annotations

import argparse
import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd

import kpi_rules
import pool as pool_mod
import receiving_log as rl
from common import (
    DIRS,
    ROOT,
    StopWithMessage,
    atomic_write_csv,
    clean_text,
    ensure_dirs,
    join_unique,
    load_config,
    mawb_match_key,
    new_run_id,
    setup_logging,
    write_json,
    write_workbook,
    write_status_check_list,
)
from dc_bridge import load_unified_master

QUERY_STATUSES = {
    "Late",
    "Late(PickUp)",
    "Late(Handover)",
    "Late(Manifest)",
    "Late(Others)",
    "No A-Scan",
    "No A-Scan (未查询)",
    "No Ascan No Track",
    "Missing Ascan",
    "51005",
    "Wrong Address",
}

FOUND_SUFFIX = "(Found)"
DEFAULT_CROSS_HUB_FOUND = {"LAX-DFW": "LAX-DFW(Found)", "EWR-ORD": "EWR-ORD(Found)"}


def cross_hub_base(tag) -> str:
    """Strip the (Found) suffix so receiving hub maps still match."""
    text = clean_text(tag) or ""
    if text.endswith(FOUND_SUFFIX):
        return text[: -len(FOUND_SUFFIX)].strip()
    return text


def mark_cross_hub_found(master: pd.DataFrame, pulled_idx: pd.Index, cfg: dict, logger):
    """Retag pool-pulled cross-hub rows as LAX-DFW(Found) / EWR-ORD(Found)."""
    mapping = cfg.get("cross_hub_found") or DEFAULT_CROSS_HUB_FOUND
    if master.empty or "跨Hub转运" not in master.columns or not mapping:
        return master
    cross = master["跨Hub转运"].astype("string")
    pulled_mask = master.index.isin(pulled_idx)
    for raw, found in mapping.items():
        hit = pulled_mask & cross.isin([raw, found])
        n = int(hit.sum())
        if n:
            master.loc[hit, "跨Hub转运"] = found
            logger.info("跨Hub转运：Pool 捞出后 %s -> %s，共 %d 条", raw, found, n)
    return master


# ------------------------------------------------------------- receiving match

def match_receiving_to_pool(active: pd.DataFrame, history: pd.DataFrame, cfg: dict, logger):
    """MAWB is the key, hub is verified. Ambiguous batches go to 待核对, never averaged."""
    empty = pd.DataFrame(columns=["TNO", "MAWB", "Hub", "reason", "candidates"])
    if active.empty or history.empty:
        logger.info("Pool 或收货历史为空，本次没有可写入的收货时间。")
        return active, empty, empty

    # Only events cleared for matching. No MAWB, non-standard number, unresolved hub
    # or unparsable time all stay in the history but never auto-match.
    usable = history[(history["match_eligible"] == "Yes")
                     & history["receiving_time"].notna()
                     & history["mawb_key"].notna()].copy()
    usable["receiving_time"] = pd.to_datetime(usable["receiving_time"], errors="coerce")
    usable = usable[usable["receiving_time"].notna()]
    held = history[history["match_eligible"] != "Yes"]
    logger.info("可用于匹配的收货事件 %d 条，覆盖 %d 个 MAWB；另有 %d 条待核对事件不参与匹配",
                len(usable), usable["mawb_key"].nunique(), len(held))
    if len(held):
        for label, count in held["match_status"].value_counts().items():
            logger.info("    不参与匹配 - %s: %d", label, count)

    expectation = cfg.get("receiving_hub_expectation", {}) or {}
    equivalence = {str(k).upper(): {str(x).upper() for x in v}
                   for k, v in (cfg.get("hub_equivalence", {}) or {}).items()}
    by_mawb = {key: sub for key, sub in usable.groupby("mawb_key")}

    active = active.copy()
    active["_mawb_key"] = active["MAWB"].map(mawb_match_key)
    pending_rows, conflict_rows = [], []
    matched = 0
    already = pd.to_datetime(active.get("receiving_time"), errors="coerce").notna()
    todo = active.loc[~already]

    for idx, row in todo.iterrows():
        key = row["_mawb_key"]
        if not key or key not in by_mawb:
            continue
        events = by_mawb[key]
        pool_hub = (clean_text(row.get("Hub")) or "").upper()
        cross = cross_hub_base(row.get("跨Hub转运"))
        allowed = {pool_hub} if pool_hub else set()
        allowed |= {str(h).upper() for h in expectation.get(cross, [])} if cross else set()
        allowed |= equivalence.get(pool_hub, set())

        ok = events[events["Hub"].astype("string").str.upper().isin(allowed)] if allowed else events.iloc[0:0]
        bad = events[~events.index.isin(ok.index)]

        if ok.empty:
            pending_rows.append({
                "TNO": row["TNO"], "MAWB": row.get("MAWB"), "Hub": row.get("Hub"),
                "reason": "待核对_收货Hub与包裹Hub不符且无已确认映射",
                "candidates": ";".join(sorted(set(bad["Hub"].dropna().astype(str)))),
            })
            continue
        if ok["receiving_record_id"].nunique() > 1:
            # No package-level or batch-level evidence exists, so neither the first
            # nor the last arrival may be assumed. Leave receiving_time empty.
            pending_rows.append({
                "TNO": row["TNO"], "MAWB": row.get("MAWB"), "Hub": row.get("Hub"),
                "reason": "待核对_同一MAWB多个收货批次，无包裹级或批次级证据，未取任何一笔时间",
                "candidates": ";".join(
                    f"{pd.Timestamp(t):%Y-%m-%d %H:%M}" for t in sorted(ok["receiving_time"].unique())
                ),
            })
            continue

        event = ok.iloc[0]
        if str(row.get("kpi_counted") or "").strip().lower() == "yes":
            conflict_rows.append({
                "TNO": row["TNO"], "MAWB": row.get("MAWB"), "Hub": row.get("Hub"),
                "reason": "已计入历史KPI后才确认实际收货",
                "candidates": f"{event['receiving_time']:%Y-%m-%d %H:%M:%S}",
            })
            continue

        active.at[idx, "receiving_time"] = event["receiving_time"]
        active.at[idx, "receiving_record_id"] = event["receiving_record_id"]
        active.at[idx, "receiving_hub"] = event["Hub"]
        active.at[idx, "hub_match_status"] = (
            "一致" if pool_hub and str(event["Hub"]).upper() == pool_hub else f"已确认映射:{cross or '配置'}"
        )
        active.at[idx, "match_status"] = "已匹配收货"
        matched += 1

    active = active.drop(columns=["_mawb_key"])
    logger.info("Pool 收货匹配：成功 %d 条，待核对 %d 条，已计入后收货冲突 %d 条",
                matched, len(pending_rows), len(conflict_rows))
    return (active,
            pd.DataFrame(pending_rows, columns=empty.columns),
            pd.DataFrame(conflict_rows, columns=empty.columns))


# --------------------------------------------------------------- pool kpi math

def refresh_pool_kpi(active: pd.DataFrame, kpi_date: str, logger):
    """Recompute uncounted pool KPI dates.

    Winit uses Manifest Time calendar D0/D1; receiving does not change Winit due.
    Hour-SLA rows still take receiving_time when present.
    """
    if active.empty:
        return active, pd.DataFrame(columns=["TNO", "kpi_date", "reason"])
    active = active.copy()
    counted = active["kpi_counted"].astype("string").str.lower().eq("yes").fillna(False)
    recv = pd.to_datetime(active["receiving_time"], errors="coerce") if "receiving_time" in active.columns else pd.Series(pd.NaT, index=active.index)
    orig = pd.to_datetime(active["start_time"], errors="coerce") if "start_time" in active.columns else pd.Series(pd.NaT, index=active.index)
    clock = recv.where(recv.notna(), orig)
    times = kpi_rules.compute_kpi_times(active, clock)
    recompute = ~counted
    for col in ("kpi_start_time", "kpi_due_date", "kpi_date"):
        active.loc[recompute, col] = times.loc[recompute, col]
    active.loc[recompute, "sla_type"] = times.loc[recompute, "sla_type"].astype("object").to_numpy()
    active.loc[recompute, "sla_hours"] = pd.to_numeric(
        times.loc[recompute, "sla_hours"], errors="coerce"
    ).astype("float64").to_numpy()
    active.loc[recompute, "kpi_deadline"] = times.loc[recompute, "kpi_deadline"].to_numpy()

    if "kpi_start_time_source" not in active.columns:
        active["kpi_start_time_source"] = pd.NA
    cal = times["sla_type"].astype("string").isin(["calendar_d0", "calendar_d1"])
    active.loc[recompute & cal, "kpi_start_time_source"] = "Original Data"
    active.loc[recompute & ~cal & recv.notna(), "kpi_start_time_source"] = "Receiving Log"
    blank_src = active["kpi_start_time_source"].isna() | active["kpi_start_time_source"].astype("string").isin(["", "<NA>"])
    active.loc[recompute & ~cal & recv.isna() & orig.notna() & blank_src, "kpi_start_time_source"] = "Original Data"

    dated = active["kpi_date"].notna()
    due_now = dated & (active["kpi_date"].astype("string") == kpi_date) & ~counted
    overdue = dated & (active["kpi_date"].astype("string") < kpi_date) & ~counted
    future = dated & (active["kpi_date"].astype("string") > kpi_date) & ~counted

    active.loc[due_now, "pool_state"] = pool_mod.STATE_PENDING
    active.loc[future | ~dated, "pool_state"] = active.loc[future | ~dated, "pool_state"].fillna(
        pool_mod.STATE_TRACKING
    )
    active.loc[future, "pool_state"] = pool_mod.STATE_TRACKING

    logger.info("Pool 到期检查（本次 KPI 日期 %s）：本次转入 %d，未来到期保留 %d，"
                "过去到期未计入 %d，尚无确认日期 %d",
                kpi_date, int(due_now.sum()), int(future.sum()),
                int(overdue.sum()), int((~dated).sum()))
    if int(overdue.sum()):
        logger.warning("有 %d 条 Pool 记录已过期但未计入，已写入异常清单，未静默塞进本次。",
                       int(overdue.sum()))
    overdue_df = active.loc[overdue, ["TNO", "kpi_date"]].copy()
    overdue_df["reason"] = f"过去到期未计入（本次KPI日期 {kpi_date}）"
    return active, overdue_df


# ----------------------------------------------------------------- master build

def select_pool_for_master(active: pd.DataFrame, _receiving_std: pd.DataFrame,
                           kpi_date: str, logger):
    """Only uncounted rows whose kpi_date is today enter the master.

    Receiving still updates the pool in place. If that makes a row due
    today, it is pulled here; if it is still in the future, it waits.
    """
    empty = pd.DataFrame(columns=active.columns if len(active) else [])
    if active.empty:
        logger.info("Pool 捞出 0 条（池为空）")
        return empty, empty
    counted = active["kpi_counted"].astype("string").str.lower().eq("yes").fillna(False)
    due_now = active["kpi_date"].astype("string").eq(kpi_date) & ~counted
    pulled, held = active.loc[due_now].copy(), active.loc[~due_now].copy()
    logger.info("Pool 捞出 %d 条（当日到期未计入）；其余 %d 条留在池中等待",
                int(due_now.sum()), int((~due_now).sum()))
    return pulled, held


def build_waiting_pool_rows(waiting: pd.DataFrame, kpi_date: str, now: str,
                            labels: dict) -> pd.DataFrame:
    """Park not-yet-due rows in Active_Pool. Do not mark them counted."""
    if waiting.empty:
        return pd.DataFrame(columns=pool_mod.POOL_COLS)
    rows = pd.DataFrame(index=waiting.index)
    for col in pool_mod.POOL_COLS:
        if col in waiting.columns:
            rows[col] = waiting[col]
    rows["TNO"] = waiting["TNO"]
    if "Source" in rows.columns:
        rows["Source"] = [join_unique([src, labels["pool"]]) for src in rows["Source"]]
    rows["last_update_time"] = now
    rows["pool_entry_date"] = rows["pool_entry_date"].fillna(kpi_date) if "pool_entry_date" in rows.columns else kpi_date
    rows["pool_state"] = pool_mod.STATE_TRACKING
    return rows


def build_step1_master(master: pd.DataFrame, active: pd.DataFrame, pulled: pd.DataFrame,
                       kpi_date: str, cfg: dict, logger):
    labels = cfg["source_labels"]
    master = master.copy()
    master["TNO"] = master["TNO"].astype("string")
    master = master.set_index("TNO", drop=False)
    three_idx = master.index

    if not active.empty:
        active = active.copy()
        active["TNO"] = active["TNO"].astype("string")
        active = active.drop_duplicates("TNO", keep="last").set_index("TNO", drop=False)
    if not pulled.empty:
        pulled = pulled.copy()
        pulled["TNO"] = pulled["TNO"].astype("string")
        pulled = pulled.drop_duplicates("TNO", keep="last").set_index("TNO", drop=False)

    full_pool_idx = active.index if not active.empty else pd.Index([], dtype="string")
    pulled_idx = pulled.index if not pulled.empty else pd.Index([], dtype="string")
    pool_only = pulled.loc[pulled_idx.difference(three_idx)] if not pulled.empty else pd.DataFrame()

    # Pool-only TNOs keep the business fields the pool preserved for them.
    if len(pool_only):
        add = pd.DataFrame(index=pool_only.index, columns=master.columns)
        for col in master.columns:
            if col in pool_only.columns:
                add[col] = pool_only[col]
        add["TNO"] = pool_only["TNO"]
        master = pd.concat([master, add])
        logger.info("Pool 独有 TNO %d 条并入主表", len(pool_only))

    in_pool = pd.Series(master.index.isin(full_pool_idx), index=master.index)
    in_three = pd.Series(master.index.isin(three_idx), index=master.index)

    # original_source keeps the pre-pool Source. Rows re-entering from the pool already
    # carry it, so do not overwrite that with a Source that has the pool label appended.
    stored_original = pd.Series(pd.NA, index=master.index, dtype="object")
    if "original_source" in master.columns:
        stored_original = master["original_source"]
    if len(pool_only) and "original_source" in pool_only.columns:
        shared = master.index.intersection(pool_only.index)
        stored_original.loc[shared] = pool_only.loc[shared, "original_source"]
    master["original_source"] = stored_original.where(
        stored_original.notna(), master["Source"]
    )
    master["Source"] = [
        join_unique([src, labels["pool"]]) if flag else src
        for src, flag in zip(master["Source"], in_pool)
    ]
    route = pd.Series(labels["route_three_sources"], index=master.index, dtype="object")
    route = route.mask(in_pool & in_three, labels["route_both"])
    route = route.mask(in_pool & ~in_three, labels["route_pool"])
    master["record_route"] = route
    entry = pd.Series(labels.get("entry_three", "当天三源"), index=master.index, dtype="object")
    entry = entry.mask(in_pool & in_three, labels.get("entry_both", "当天三源+Pool"))
    entry = entry.mask(in_pool & ~in_three, labels.get("entry_pool", "Pool捞出"))
    master["进表来源"] = entry
    logger.info("进表来源：当天三源 %d，Pool捞出 %d，当天三源+Pool %d",
                int((entry == labels.get("entry_three", "当天三源")).sum()),
                int((entry == labels.get("entry_pool", "Pool捞出")).sum()),
                int((entry == labels.get("entry_both", "当天三源+Pool")).sum()))
    mark_cross_hub_found(master, pulled_idx, cfg, logger)
    master, n_fbt_iah = kpi_rules.apply_hub_corrections(master)
    if n_fbt_iah:
        logger.info("TT FBT 无 DFW：Hub DFW -> IAH，共 %d 条", n_fbt_iah)

    # Pool-confirmed values fill in; a blank pool value never clears a good master value.
    for col in ("receiving_time", "receiving_record_id", "kpi_counted",
                "counted_kpi_date", "counted_run_id", "pool_state", "pool_entry_date"):
        master[col] = pd.NA
        if not active.empty and col in active.columns:
            shared = master.index.intersection(active.index)
            incoming = active.loc[shared, col]
            master.loc[shared, col] = incoming.where(incoming.notna(), master.loc[shared, col])

    pool_start = pd.Series(pd.NaT, index=master.index)
    if not active.empty:
        shared = master.index.intersection(active.index)
        pool_start.loc[shared] = pd.to_datetime(active.loc[shared, "kpi_start_time"], errors="coerce")

    original_start = pd.to_datetime(master["start_time"], errors="coerce")
    kpi_start = pool_start.where(pool_start.notna(), original_start)
    source = pd.Series(pd.NA, index=master.index, dtype="object")
    source = source.mask(pool_start.notna(), "Receiving Log")
    source = source.mask(pool_start.isna() & original_start.notna(), "Original Data")

    tz_bad = kpi_rules.timezone_exceptions(master["start_time"].astype("string"))
    if bool(tz_bad.any()):
        logger.warning("有 %d 行时间带时区标记，已标为异常且未做任何换算", int(tz_bad.sum()))

    times = kpi_rules.compute_kpi_times(master, kpi_start)
    cal = times["sla_type"].astype("string").isin(["calendar_d0", "calendar_d1"])
    master["kpi_start_time"] = times["kpi_start_time"]
    master["kpi_start_time_source"] = source
    master.loc[cal, "kpi_start_time_source"] = "Original Data"
    master["kpi_due_date"] = times["kpi_due_date"]
    master["kpi_date"] = times["kpi_date"]
    master["kpi_sla_hours"] = times["sla_hours"]
    master["sla_hours"] = times["sla_hours"]
    master["sla_type"] = times["sla_type"]
    master["kpi_deadline"] = times["kpi_deadline"]

    targets = cfg.get("sla_targets") or {}
    winit_cfg = targets.get("Winit") or {}
    master["sla_customer_target"] = pd.NA
    master["sla_internal_target"] = pd.NA
    master.loc[cal, "sla_customer_target"] = winit_cfg.get("customer")
    master.loc[cal, "sla_internal_target"] = winit_cfg.get("internal")

    gap = (pd.to_datetime(master["receiving_time"], errors="coerce") - original_start)
    master["start_time_gap_hours"] = (gap / pd.Timedelta(hours=1)).round(4).where(~tz_bad)

    counted = master["kpi_counted"].astype("string").str.lower().eq("yes").fillna(False)
    due_today = (master["kpi_date"].astype("string") == kpi_date) & ~counted
    excl_ord = kpi_rules.aliexpress_ord_ineligible(master)
    excl_mia = kpi_rules.aliexpress_no_ags_mia(master)
    excl_pc = kpi_rules.is_procarrier_customer(master)
    master["kpi_eligible"] = kpi_rules.apply_eligibility_exclusions(due_today, master)
    n_ord = int((due_today & excl_ord).sum())
    n_mia = int((due_today & excl_mia).sum())
    n_pc = int((due_today & excl_pc).sum())
    if n_ord:
        logger.info(
            "AliExpress（业务名称或 CUSTOMER）且 Hub=ORD：%d 条 kpi_eligible=False，留在主表但不进分母",
            n_ord,
        )
    if n_mia:
        logger.info("AliExpress 非AGS 且 Hub 为 MIA：%d 条 kpi_eligible=False，留在主表但不进分母", n_mia)
    if n_pc:
        logger.info("CUSTOMER=ProCarrier：%d 条 kpi_eligible=False（Pro Carrier T11 仍计入）", n_pc)
    master["kpi_time_exception"] = pd.NA
    master.loc[tz_bad, "kpi_time_exception"] = "时间带时区标记，未换算，请人工确认基准"
    master.loc[master["kpi_start_time"].isna(), "kpi_time_exception"] = "无可用起算时间，无法计算 kpi_due_date"
    master.loc[cal & master["kpi_due_date"].isna(), "kpi_time_exception"] = (
        "Winit 缺少有效 Manifested Time，无法计算应履约日期"
    )

    reasons = cfg["query_reason"]
    status = master["status"].astype("string")
    needs_query = status.isin(QUERY_STATUSES)
    query_reason = pd.Series(pd.NA, index=master.index, dtype="object")
    query_reason = query_reason.mask(needs_query & master["kpi_eligible"], reasons["current_kpi"])
    query_reason = query_reason.mask(
        needs_query & ~master["kpi_eligible"] & in_pool, reasons["history_tracking"]
    )
    master["query_reason"] = query_reason

    waiting = master.loc[~due_today].copy()
    master = master.loc[due_today].copy()
    logger.info(
        "Step1 主表保留本次到期 %d 行（kpi_eligible=%d）；未到期 %d 行写入 Active_Pool 等待",
        len(master), int(master["kpi_eligible"].fillna(False).astype(bool).sum()), len(waiting),
    )
    if len(master) and len(master) != master["TNO"].nunique():
        raise StopWithMessage("Step1 主表出现重复 TNO，已中止。")
    return master.reset_index(drop=True), waiting.reset_index(drop=True)


# ------------------------------------------------------------------------ main

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="第一步：整合并生成查询清单")
    parser.add_argument("--kpi-date", required=True, help="本次 KPI 日期 YYYY-MM-DD（必须显式指定）")
    parser.add_argument("--merge", required=True, choices=["run", "reuse"],
                        help="run=重新跑三源合并；reuse=使用本次已生成的宽表")
    parser.add_argument("--allow-stale-master", action="store_true",
                        help="明确确认使用超过时效阈值的旧宽表")
    parser.add_argument("--receiving-date", default=None,
                        help="收货文件内没有日期列时，显式提供收货日期 YYYY-MM-DD")
    args = parser.parse_args(argv)

    try:
        datetime.strptime(args.kpi_date, "%Y-%m-%d")
    except ValueError:
        print(f"KPI 日期格式错误：{args.kpi_date}，应为 YYYY-MM-DD")
        return 2
    external_date = None
    if args.receiving_date:
        try:
            external_date = datetime.strptime(args.receiving_date, "%Y-%m-%d").date()
        except ValueError:
            print(f"收货日期格式错误：{args.receiving_date}，应为 YYYY-MM-DD")
            return 2

    ensure_dirs()
    run_id = new_run_id("step1")
    logger = setup_logging(run_id, "step1")
    work = DIRS["working"] / run_id
    out_dir = DIRS["step1_output"] / run_id
    work.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        cfg = load_config("closed_loop.yaml")
        logger.info("run_id=%s  本次 KPI 日期=%s  合并模式=%s", run_id, args.kpi_date, args.merge)

        master, meta = load_unified_master(cfg, args.merge, args.allow_stale_master, logger)
        active, closed, first_time = pool_mod.load_pool(logger)

        recv_dir = DIRS["receiving_input"]
        recv_files = [
            p for p in recv_dir.iterdir()
            if p.is_file() and not p.name.startswith("~$")
            and p.suffix.lower() in {".xlsx", ".xls", ".csv"}
        ] if recv_dir.exists() else []
        empty_recv = pd.DataFrame(columns=["TNO", "MAWB", "Hub", "reason", "candidates"])
        if not recv_files:
            logger.info("Receiving_Log/Input 为空，跳过收货匹配（不影响 KPI / Pool）")
            std = pd.DataFrame()
            used_files = []
            history = pd.DataFrame()
            pending, already_counted = empty_recv, empty_recv
        else:
            rl_cfg = load_config("receiving_log.yaml")
            std, used_files = rl.standardize_inputs(rl_cfg, run_id, external_date, logger)
            history = rl.append_history(std, logger)
            active, pending, already_counted = match_receiving_to_pool(active, history, cfg, logger)
        active, overdue = refresh_pool_kpi(active, args.kpi_date, logger)
        pulled, pool_held = select_pool_for_master(active, std, args.kpi_date, logger)

        step1, waiting = build_step1_master(master, active, pulled, args.kpi_date, cfg, logger)
        before_due = len(step1)
        step1, removed_due = kpi_rules.split_removed_rows(step1)
        waiting, removed_wait = kpi_rules.split_removed_rows(waiting)
        if len(removed_due):
            removed_due["删除范围"] = "本次到期"
        if len(removed_wait):
            removed_wait["删除范围"] = "未到期，不入 Pool"
        removed = pd.concat([removed_due, removed_wait], ignore_index=True)
        removed.to_parquet(work / "removed_rows.parquet", index=False)
        if len(removed):
            for reason, n in removed["删除原因"].value_counts().items():
                logger.info("删除 %s：%d", reason, int(n))
        logger.info(
            "删除合计 %d（本次到期 %d，未到期不入池 %d）；本次到期主表 %d -> %d",
            len(removed), len(removed_due), len(removed_wait), before_due, len(step1),
        )
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        wait_rows = build_waiting_pool_rows(waiting, args.kpi_date, now, cfg["source_labels"])
        if len(wait_rows):
            logger.info("未到期写入 Active_Pool %d 条（其中已有 A-scan %d）",
                        len(wait_rows),
                        int(pd.to_datetime(wait_rows.get("first_ascan_time"), errors="coerce").notna().sum())
                        if "first_ascan_time" in wait_rows.columns else 0)
            active = pool_mod.upsert_active(active, wait_rows, logger)

        # ---- outputs
        logger.info("写 Working 中间数据")
        step1.to_parquet(work / "unified_master_step1.parquet", index=False)
        atomic_write_csv(work / "pool_waiting_not_due.csv", waiting)
        held = history[history["match_eligible"] != "Yes"] if len(history) else history
        for name, frame in (
            ("receiving_standardized.csv", std),
            ("receiving_pending_review.csv", pending),
            ("receiving_not_matchable.csv", held),
            ("already_counted_receiving_conflict.csv", already_counted),
            ("pool_overdue_not_counted.csv", overdue),
            ("pool_not_pulled.csv", pool_held),
        ):
            atomic_write_csv(work / name, frame)

        query = step1[step1["query_reason"].notna()].drop_duplicates("TNO")

        logger.info("写 Step1 输出（主表 %d 行 / Status Check 清单 %d 行）", len(step1), len(query))
        font = cfg["report"].get("font_name")
        write_workbook(out_dir / "Unified_Master_Step1.xlsx",
                       {"Unified_Master_Step1": step1}, font=font)
        list_path = out_dir / "Status_Check_List.xlsx"
        write_status_check_list(list_path, query["TNO"], font=font)
        logger.info("Status Check 单号清单（仅 Tracking Number）：%s",
                    list_path.relative_to(ROOT))

        pool_mod.save_pool(active, closed, logger, cfg, reason=f"step1 {run_id}")

        # ---- archive the receiving inputs actually used
        archived = []
        if used_files:
            dest = DIRS["archive"] / "Receiving_Log" / run_id
            dest.mkdir(parents=True, exist_ok=True)
            for path in used_files:
                shutil.move(str(path), str(dest / path.name))
                archived.append(path.name)
            logger.info("已归档本次使用的收货文件 %d 个 -> %s", len(archived),
                        dest.relative_to(ROOT))

        pending_date_files = sorted(
            std.loc[std["receiving_date_status"] != rl.DATE_CONFIRMED, "source_file"]
            .dropna().unique().tolist()
        ) if len(std) else []
        if pending_date_files:
            logger.warning("=" * 70)
            logger.warning("本批次收货日期待确认：%s", "、".join(pending_date_files))
            logger.warning("日期取自运行参数 --receiving-date %s，不是文件内容。",
                           args.receiving_date)
            logger.warning("确认或更正前，第二步会拒绝出正式 KPI。")
            logger.warning("确认：python src\\receiving_tool.py confirm --source-file \"<文件名>\"")
            logger.warning("更正：python src\\receiving_tool.py correct --source-file \"<文件名>\" "
                           "--new-date YYYY-MM-DD")
            logger.warning("=" * 70)

        manifest = {
            "run_id": run_id,
            "stage": "step1",
            "state": "awaiting_status_check",
            "kpi_date": args.kpi_date,
            "receiving_date_argument": args.receiving_date,
            "receiving_date_pending_files": pending_date_files,
            "receiving_date_confirmed": not pending_date_files,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "unified_master": meta,
            "pool_first_time": first_time,
            "counts": {
                "master_rows": int(len(step1)),
                "unique_tno": int(step1["TNO"].nunique()),
                "kpi_eligible": int(step1["kpi_eligible"].sum()),
                "query_rows": int(len(query)),
                "query_current_kpi": int((query["query_reason"] == cfg["query_reason"]["current_kpi"]).sum()),
                "query_history": int((query["query_reason"] == cfg["query_reason"]["history_tracking"]).sum()),
                "master_rows_total": int(len(step1)),
                "receiving_records": int(len(std)),
                "receiving_not_matchable": int(len(held)),
                "receiving_pending_review": int(len(pending)),
                "pool_pending_kpi": int((active["pool_state"] == pool_mod.STATE_PENDING).sum()) if len(active) else 0,
                "pool_overdue_not_counted": int(len(overdue)),
                "pool_pulled": int(len(pulled)),
                "pool_not_pulled": int(len(pool_held)),
                "pool_waiting_not_due": int(len(waiting)),
                "already_counted_receiving_conflict": int(len(already_counted)),
                "removed_rows": int(len(removed)),
                "removed_due": int(len(removed_due)),
                "removed_not_due": int(len(removed_wait)),
            },
            "archived_receiving_files": archived,
        }
        write_json(work / "manifest.json", manifest)
        write_json(out_dir / "manifest.json", manifest)

        logger.info("=" * 70)
        logger.info("第一步完成。run_id=%s", run_id)
        logger.info("查询清单：%s", (out_dir / "Status_Check_List.xlsx").relative_to(ROOT))
        logger.info("请用 Status_Check_List.xlsx 去做系统查询（只有 Tracking Number 一列）。")
        logger.info("把下载结果放到 Status_Check/Input，然后运行第二步。")
        logger.info("=" * 70)
        return 0
    except StopWithMessage as exc:
        logger.error("已停止：%s", exc)
        return 1
    except Exception:
        logger.error("未处理错误：\n%s", traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
