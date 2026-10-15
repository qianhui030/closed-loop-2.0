# -*- coding: utf-8 -*-
"""Step 2: apply the status check, publish the KPI report, then maintain the pool."""
from __future__ import annotations

import argparse
import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import kpi_rules
import pivots
import pool as pool_mod
from common import (
    DIRS,
    ROOT,
    StopWithMessage,
    atomic_write_csv,
    clean_text,
    clean_tno,
    ensure_dirs,
    find_header_row,
    load_config,
    new_run_id,
    parse_dt,
    read_json,
    resolve_columns,
    setup_logging,
    write_json,
)

REQUIRED_APPLY_FIELDS = ("tno", "ascan_time", "last_event_code", "last_event_desc")
OPTIONAL_SC_FIELDS = (
    "last_event_time",
    "last_status_event_desc",
    "last_attempt_event_desc",
    "last_public_status_event_desc",
    "arrived_sc_date_57407",
)
SC_TEXT_FIELDS = {
    "last_event_code": "sc_event_code",
    "last_event_desc": "sc_event_desc",
    "last_status_event_desc": "sc_status_event_desc",
    "last_attempt_event_desc": "sc_attempt_event_desc",
    "last_public_status_event_desc": "sc_public_status_event_desc",
    "arrived_sc_date_57407": "sc_arrived_sc_date_57407",
}


def receiving_date_ready(batch: dict) -> tuple[bool, str]:
    pending = batch.get("receiving_date_pending_files") or []
    if batch.get("receiving_date_confirmed") is True and not pending:
        return True, "收货日期已确认"
    if pending:
        return False, (
            "本批次收货日期待确认：" + "、".join(pending)
            + "。请先运行 receiving_tool.py confirm 或 correct，再出正式 KPI。"
        )
    if batch.get("receiving_date_argument") and batch.get("receiving_date_confirmed") is not True:
        return False, (
            f"本批次收货日期 {batch.get('receiving_date_argument')} 仍待确认，"
            "不能作为正式 KPI。"
        )
    return True, "本批次没有依赖运行参数的收货日期"


# ------------------------------------------------------------------ batch pick

def list_pending(logger) -> list[dict]:
    out = []
    for path in sorted(DIRS["working"].glob("*/manifest.json")):
        try:
            data = read_json(path)
        except Exception:  # noqa: BLE001
            continue
        if data.get("stage") == "step1" and data.get("state") == "awaiting_status_check":
            out.append(data)
    return out


def pick_batch(run_id: str | None, logger) -> dict:
    pending = list_pending(logger)
    if run_id:
        for item in pending:
            if item["run_id"] == run_id:
                return item
        done = DIRS["working"] / run_id / "manifest.json"
        if done.exists():
            data = read_json(done)
            raise StopWithMessage(
                f"批次 {run_id} 的状态是 {data.get('state')}，不是待处理。"
                "重复运行不会重复计数。如需重做，请先新建第一步批次。"
            )
        raise StopWithMessage(f"找不到批次 {run_id}。")
    if not pending:
        raise StopWithMessage(
            "没有待处理的第一步批次。请先运行第一步。"
        )
    if len(pending) > 1:
        lines = "\n".join(
            f"  --run-id {p['run_id']}   KPI日期={p['kpi_date']}  生成于 {p['created_at']}"
            for p in sorted(pending, key=lambda x: x["created_at"])
        )
        raise StopWithMessage(
            f"有 {len(pending)} 个待处理批次，请明确指定一个，不自动挑选：\n{lines}"
        )
    return pending[0]


# ------------------------------------------------------------- status check io

def read_status_check(cfg: dict, logger) -> tuple[pd.DataFrame, list[Path]]:
    files = [p for p in sorted(DIRS["status_input"].iterdir())
             if p.is_file() and not p.name.startswith("~$")] if DIRS["status_input"].exists() else []
    if not files:
        raise StopWithMessage(
            f"Status_Check/Input 为空。请先把系统查询结果放进去：{DIRS['status_input']}"
        )
    scan = int(cfg.get("header_scan_rows", 6))
    frames = []
    for path in files:
        if path.suffix.lower() in {".xlsx", ".xlsm"}:
            try:
                from python_calamine import CalamineWorkbook
                cwb = CalamineWorkbook.from_path(str(path))
                rows = [
                    list(r) for r in cwb.get_sheet_by_index(0).to_python(skip_empty_area=False)
                ]
            except Exception:
                wb = load_workbook(path, read_only=True, data_only=True)
                rows = [list(r) for r in wb[wb.sheetnames[0]].iter_rows(values_only=True)]
                wb.close()
        else:
            rows = None
            for enc in ("utf-8-sig", "utf-8", "gbk", "cp1252", "latin1"):
                try:
                    rows = pd.read_csv(path, dtype=str, header=None, encoding=enc).values.tolist()
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            if rows is None:
                raise StopWithMessage(f"无法解码查询结果文件：{path.name}")
        if not rows:
            raise StopWithMessage(f"查询结果文件为空：{path.name}")
        head = find_header_row(rows, cfg["columns"], scan)
        frame = pd.DataFrame(rows[head + 1:], columns=[clean_text(c) or f"col_{i}"
                                                       for i, c in enumerate(rows[head])])
        frame["__source_file"] = path.name
        frames.append(frame)
        logger.info("  读取查询结果 %s（%d 行）", path.name, len(frame))

    raw = pd.concat(frames, ignore_index=True)
    cols = resolve_columns([c for c in raw.columns if not c.startswith("__")], cfg["columns"])
    logger.info("查询结果命中列：%s", cols)
    if "tno" not in cols:
        raise StopWithMessage(
            f"查询结果里找不到单号列。实际列：{list(raw.columns)}\n"
            "请在 config/status_check.yaml 的 columns.tno 里补上真实列名。"
        )
    out = pd.DataFrame({
        "TNO": raw[cols["tno"]].map(clean_tno),
        "sc_ascan_time": raw[cols["ascan_time"]].map(
            lambda v: parse_dt(v, cfg.get("datetime_formats"))) if "ascan_time" in cols else pd.NaT,
        "sc_source_file": raw["__source_file"],
    })
    for logical, dest in SC_TEXT_FIELDS.items():
        if logical in cols:
            out[dest] = raw[cols[logical]].map(clean_text)
        else:
            out[dest] = pd.NA
    if "last_event_time" in cols:
        out["sc_event_time"] = raw[cols["last_event_time"]].map(
            lambda v: parse_dt(v, cfg.get("datetime_formats"))
        )
    else:
        out["sc_event_time"] = pd.NaT
    missing_fields = [k for k in (*REQUIRED_APPLY_FIELDS[1:], *OPTIONAL_SC_FIELDS) if k not in cols]
    if missing_fields:
        logger.warning("查询结果缺少这些列，相关判定会退化：%s", missing_fields)
    out = out[out["TNO"].notna()].copy()
    stats = {
        "raw_rows": int(len(raw)),
        "rows_with_tno": int(len(out)),
        "duplicate_tnos": sorted(out.loc[out["TNO"].duplicated(), "TNO"].unique().tolist()),
        "missing_logical_columns": missing_fields,
        "files": [p.name for p in files],
    }
    if stats["duplicate_tnos"]:
        logger.warning("查询结果有 %d 个重复 TNO，保留信息最全的一条",
                       len(stats["duplicate_tnos"]))
        out["_score"] = out["sc_ascan_time"].notna().astype(int) + out["sc_event_code"].notna().astype(int)
        out = out.sort_values("_score", ascending=False).drop_duplicates("TNO").drop(columns="_score")
    return out, files, stats


# ------------------------------------------------------------------ status calc

def report_coverage(master: pd.DataFrame, sc: pd.DataFrame, logger) -> dict:
    """Coverage is measured against Status_Check_List, not against the whole master."""
    ids = set(master["TNO"].astype("string"))
    query = master[master["query_reason"].notna()].drop_duplicates("TNO")
    query_ids = set(query["TNO"].astype("string"))
    sc_ids = set(sc["TNO"].astype("string"))
    hit, miss = query_ids & sc_ids, query_ids - sc_ids
    extra = sc_ids - query_ids
    denom = max(len(query_ids), 1)
    logger.info("=== 查询覆盖率（分母 = Status_Check_List 唯一 TNO）===")
    logger.info("  完整主表 TNO           ： %8d  （对照，不是分母）", len(ids))
    logger.info("  应查询数量             ： %8d", len(query_ids))
    logger.info("  匹配成功数量           ： %8d  (%.2f%%)", len(hit), 100 * len(hit) / denom)
    logger.info("  缺失数量               ： %8d  (%.2f%%)", len(miss), 100 * len(miss) / denom)
    logger.info("  多余数量               ： %8d（其中不在主表 %d）",
                len(extra), len(extra - ids))
    return {
        "master_tno_total": len(ids),
        "should_query": len(query_ids),
        "matched": len(hit),
        "missing": len(miss),
        "extra": len(extra),
        "extra_not_in_master": len(extra - ids),
        "missing_tnos_sample": sorted(miss)[:20],
        "extra_tnos_sample": sorted(extra)[:20],
    }


def conflict_preview(master: pd.DataFrame, sc: pd.DataFrame, logger, limit: int = 8) -> dict:
    """Count original vs query A-scan differences. Does not change any time."""
    query = master[master["query_reason"].notna()].drop_duplicates("TNO")
    merged = query.merge(sc, on="TNO", how="inner")
    original = pd.to_datetime(merged["first_ascan_time"], errors="coerce")
    from_sc = pd.to_datetime(merged["sc_ascan_time"], errors="coerce")
    conflict = original.notna() & from_sc.notna() & (original != from_sc)
    n = int(conflict.sum())
    logger.info("  原 A-scan 与查询 A-scan 都有效但不同：%d 条（仅审计，不改时间优先级）", n)
    samples = []
    if n:
        cols = [c for c in ("TNO", "first_ascan_time", "sc_ascan_time") if c in merged.columns]
        preview = merged.loc[conflict, cols].head(limit)
        for _, row in preview.iterrows():
            sample = {
                "TNO": str(row["TNO"]),
                "original": None if pd.isna(row["first_ascan_time"]) else str(row["first_ascan_time"]),
                "status_check": None if pd.isna(row["sc_ascan_time"]) else str(row["sc_ascan_time"]),
            }
            samples.append(sample)
            logger.info("    对照  TNO=%s  原=%s  查询=%s",
                        sample["TNO"], sample["original"], sample["status_check"])
    return {"conflict_both_valid": n, "conflict_samples": samples}


def apply_status(master: pd.DataFrame, sc: pd.DataFrame, sc_cfg: dict, logger):
    master = master.copy()
    master["TNO"] = master["TNO"].astype("string")
    merged = master.merge(sc, on="TNO", how="left", indicator="_sc")
    queried = merged["_sc"].eq("both")
    logger.info("主表 %d 行中有查询结果 %d 行；无查询结果 %d 行（全部保留，不从分母消失）",
                len(merged), int(queried.sum()), int((~queried).sum()))

    original = pd.to_datetime(merged["first_ascan_time"], errors="coerce")
    from_sc = pd.to_datetime(merged["sc_ascan_time"], errors="coerce")
    conflict = original.notna() & from_sc.notna() & (original != from_sc)
    policy = sc_cfg.get("ascan_conflict_policy", "keep_original")
    if policy == "prefer_status_check":
        final_ascan = from_sc.where(from_sc.notna(), original)
    else:
        final_ascan = original.where(original.notna(), from_sc)
    merged["ascan_conflict"] = conflict.map({True: "Yes", False: None})
    merged["ascan_source_final"] = pd.NA
    merged.loc[final_ascan.notna() & final_ascan.eq(original), "ascan_source_final"] = "Original"
    merged.loc[final_ascan.notna() & final_ascan.eq(from_sc) & original.isna(), "ascan_source_final"] = "Status Check"
    merged["original_ascan_time"] = original
    merged["final_ascan_time"] = final_ascan
    merged["first_ascan_time"] = final_ascan
    if int(conflict.sum()):
        logger.warning("原 A-scan 与查询 A-scan 冲突 %d 条，按配置 %s 处理并写核对清单",
                       int(conflict.sum()), policy)

    times = kpi_rules.compute_kpi_times(merged, merged["kpi_start_time"])
    merged["sla_type"] = times["sla_type"]
    merged["kpi_deadline"] = times["kpi_deadline"]
    merged["kpi_due_date"] = times["kpi_due_date"]
    merged["kpi_sla_hours"] = times["sla_hours"]
    merged["sla_hours"] = times["sla_hours"]
    cal = times["sla_type"].astype("string").isin(["calendar_d0", "calendar_d1"])
    merged.loc[cal, "kpi_start_time"] = times.loc[cal, "kpi_start_time"]
    time_diff = kpi_rules.time_diff_hours(final_ascan, merged["kpi_start_time"])
    merged["kpi_time_diff_hours"] = time_diff
    merged["time_diff_hours"] = time_diff

    base = kpi_rules.classify_base_status(
        final_ascan, time_diff, times["sla_hours"],
        sla_type=times["sla_type"],
        kpi_due_date=times["kpi_due_date"],
        kpi_deadline=times["kpi_deadline"],
    )
    status = kpi_rules.subclass_no_ascan(base, queried, merged["sc_event_code"], sc_cfg)
    status = kpi_rules.split_late(status, kpi_rules.start_type_series(merged), sc_cfg)
    wa_fields = {
        "last_event_code": merged["sc_event_code"],
        "last_event_desc": merged["sc_event_desc"],
        "last_status_event_desc": merged["sc_status_event_desc"],
        "last_attempt_event_desc": merged["sc_attempt_event_desc"],
        "last_public_status_event_desc": merged["sc_public_status_event_desc"],
        "arrived_sc_date_57407": merged["sc_arrived_sc_date_57407"],
    }
    wrong, wa_source = kpi_rules.wrong_address_match(wa_fields, sc_cfg)
    status, blocked = kpi_rules.apply_wrong_address(status, wrong, sc_cfg)

    undecidable = status.isna()
    merged["kpi_status"] = status
    merged["status"] = status
    merged["status_exception"] = pd.NA
    merged.loc[undecidable, "status_exception"] = (
        "无法判断：有 A-scan 但无法计算应履约日期或小时制 SLA"
    )
    merged.loc[blocked, "status_exception"] = "命中 Wrong Address 条件但当前为 On Time，未自动覆盖"
    merged["wrong_address_hit"] = wrong.map({True: "Yes", False: None})
    merged["wrong_address_source"] = wa_source.where(wrong, pd.NA)
    merged["onTime_count"] = kpi_rules.on_time_count(status)
    merged["queried"] = queried.map({True: "Yes", False: "No"})
    if int(wrong.sum()):
        logger.info("Wrong Address 命中 %d 条（其中 On Time 未覆盖 %d）",
                    int(wrong.sum()), int(blocked.sum()))

    if int(undecidable.sum()):
        logger.warning("有 %d 条无法判断状态，已写异常清单，未错误标为 On Time", int(undecidable.sum()))
    return merged.drop(columns=["_sc"])


def still_tracking(updated: pd.DataFrame) -> pd.Series:
    """Keep in Active_Pool: no A-scan yet, or A-scan present but this KPI date is still ahead."""
    has_ascan = updated["final_ascan_time"].notna()
    eligible = kpi_rules.coerce_eligible(updated["kpi_eligible"])
    counted = updated["kpi_counted"].astype("string").str.lower().eq("yes").fillna(False)
    return ~has_ascan | (~eligible & ~counted)


def finished_tracking(updated: pd.DataFrame) -> pd.Series:
    """Leave the pool: A-scan is in and the due KPI has been (or is being) counted."""
    has_ascan = updated["final_ascan_time"].notna()
    eligible = kpi_rules.coerce_eligible(updated["kpi_eligible"])
    counted = updated["kpi_counted"].astype("string").str.lower().eq("yes").fillna(False)
    return has_ascan & (eligible | counted)


def publish_kpi_frames(updated: pd.DataFrame, sc_cfg: dict, logger):
    """Keep ineligible rows on the report; official counts use kpi_eligible only."""
    updated = updated.copy()
    updated, n_fbt_iah = kpi_rules.apply_hub_corrections(updated)
    if n_fbt_iah:
        logger.info("TT FBT 无 DFW：Hub DFW -> IAH，共 %d 条", n_fbt_iah)
    updated["kpi_eligible"] = kpi_rules.apply_eligibility_exclusions(
        kpi_rules.coerce_eligible(updated["kpi_eligible"]), updated
    )
    if "kpi_status" in updated.columns:
        updated["kpi_status"] = kpi_rules.split_late(
            updated["kpi_status"], kpi_rules.start_type_series(updated), sc_cfg
        )
        updated["status"] = updated["kpi_status"]
        if "onTime_count" in updated.columns:
            updated["onTime_count"] = kpi_rules.on_time_count(updated["kpi_status"])
    eligible = updated[updated["kpi_eligible"]].drop_duplicates("TNO").copy()
    logger.info("本次到期主表 TNO：%d", int(updated["TNO"].nunique()))
    logger.info("本次 KPI 分母（kpi_eligible 唯一 TNO）：%d", len(eligible))
    if eligible.empty:
        return eligible, updated, {}, 0.0

    counts = eligible["kpi_status"].fillna("(未分类)").value_counts().to_dict()
    rate = float(eligible["onTime_count"].mean()) if len(eligible) else 0.0
    logger.info("本次状态分布：%s", counts)
    logger.info("本次总体 On Time Rate（按包裹数加权）= %.4f%%", rate * 100)
    unclassified = eligible["kpi_status"].isna().sum() + int(
        (eligible["kpi_status"].astype("string") == sc_cfg["no_ascan_subclass"]["not_queried"]).sum()
    )
    if unclassified:
        logger.warning("仍有 %d 条未完成分类的 No A-scan，已保留在分母并在报表中显示。", unclassified)

    report = updated.copy()
    report["status"] = report["kpi_status"].fillna("(未分类)")
    report["Hub"] = report["Hub"].fillna("(空Hub)")
    report["CUSTOMER"] = report["CUSTOMER"].fillna("(空客户)")
    report["业务名称"] = report["业务名称"].fillna("(空业务)")
    report["跨Hub转运"] = report["跨Hub转运"].fillna("非跨Hub")
    report["进表来源"] = report["进表来源"].fillna("当天三源")
    report["start_time_type"] = report["start_time_type"].fillna("(空)") if "start_time_type" in report.columns else "(空)"
    report["kpi_eligible"] = kpi_rules.eligible_label(report["kpi_eligible"])
    return eligible, report, counts, rate


def report_sheets(report: pd.DataFrame, exceptions: pd.DataFrame, outside: pd.DataFrame,
                  eligible: pd.DataFrame, kpi_date: str, cfg: dict,
                  removed_summary: pd.DataFrame | None = None,
                  removed_detail: pd.DataFrame | None = None) -> dict:
    sheets = {
        "Updated Unified Master": report,
        "删除说明": removed_summary if removed_summary is not None else kpi_rules.removal_summary(
            pd.DataFrame(columns=["删除原因"]), len(report), len(report), len(eligible)
        ),
        "本次删除": removed_detail if removed_detail is not None else kpi_rules.removal_detail(
            pd.DataFrame()
        ),
        "异常核对": exceptions,
        "非本批次查询结果": outside,
    }
    return sheets


def build_pool_rows(keep: pd.DataFrame, kpi_date: str, run_id: str, now: str) -> pd.DataFrame:
    if keep.empty:
        return pd.DataFrame(columns=pool_mod.POOL_COLS)
    rows = pd.DataFrame(index=keep.index)
    for col in pool_mod.POOL_COLS:
        if col in keep.columns:
            rows[col] = keep[col]
    rows["TNO"] = keep["TNO"]
    rows["status"] = keep["kpi_status"] if "kpi_status" in keep.columns else keep.get("status")
    rows["last_update_time"] = now
    if "pool_entry_date" in keep.columns:
        rows["pool_entry_date"] = keep["pool_entry_date"]
    rows["pool_entry_date"] = rows["pool_entry_date"].fillna(kpi_date)
    was_eligible = kpi_rules.coerce_eligible(keep["kpi_eligible"])
    rows["kpi_counted"] = keep["kpi_counted"] if "kpi_counted" in keep.columns else pd.NA
    rows.loc[was_eligible, "kpi_counted"] = "Yes"
    rows.loc[was_eligible, "counted_kpi_date"] = kpi_date
    rows.loc[was_eligible, "counted_run_id"] = run_id
    rows["pool_state"] = pool_mod.STATE_TRACKING
    return rows


def rebuild_completed_report(args, cfg: dict, sc_cfg: dict, logger, run_id2: str) -> int:
    """Re-apply kpi_eligible and rewrite the workbook. Never touches Pool or archives."""
    if not args.run_id:
        raise StopWithMessage("重出报表必须指定 --run-id")
    work = DIRS["working"] / args.run_id
    man_path = work / "manifest.json"
    if not man_path.exists():
        raise StopWithMessage(f"找不到批次 {args.run_id}")
    batch = read_json(man_path)
    if batch.get("state") != "completed":
        raise StopWithMessage(
            f"批次 {args.run_id} 状态是 {batch.get('state')}，只有已完成批次才能重出报表。"
        )
    step2 = work / "unified_master_step2.parquet"
    if not step2.exists():
        raise StopWithMessage(f"找不到 {step2.name}，无法重出报表。")
    kpi_date = batch["kpi_date"]
    logger.info("重出报表 run_id=%s（KPI 日期 %s），不改 Pool、不归档", args.run_id, kpi_date)

    updated = pd.read_parquet(step2)
    if "kpi_status" not in updated.columns and "status" in updated.columns:
        updated["kpi_status"] = updated["status"]
    if "onTime_count" not in updated.columns:
        updated["onTime_count"] = kpi_rules.on_time_count(
            updated.get("kpi_status", updated.get("status"))
        )

    saved_removed = work / "removed_rows.parquet"
    if saved_removed.exists():
        removed_prev = pd.read_parquet(saved_removed)
        logger.info("加载历史已删除记录 %d 条", len(removed_prev))
    else:
        removed_prev = pd.DataFrame(columns=["删除原因"])
    updated, removed_now = kpi_rules.split_removed_rows(updated)
    if len(removed_now):
        removed_now["删除范围"] = "本次到期"
        logger.info("重出报表新增命中删除规则 %d 条", len(removed_now))
    if len(removed_prev) and len(removed_now):
        removed = pd.concat([removed_prev, removed_now], ignore_index=True)
        if "TNO" in removed.columns:
            removed = removed.drop_duplicates("TNO", keep="last")
        removed.to_parquet(saved_removed, index=False)
    elif len(removed_now):
        removed = removed_now
        removed.to_parquet(saved_removed, index=False)
    else:
        removed = removed_prev

    eligible, report, counts, rate = publish_kpi_frames(updated, sc_cfg, logger)
    if eligible.empty:
        raise StopWithMessage(f"本次 KPI 范围为空（kpi_date={kpi_date}）。")

    removed_due = (
        removed.loc[removed["删除范围"].eq("本次到期")]
        if len(removed) and "删除范围" in removed.columns
        else removed
    )
    before_rows = len(report) + int(len(removed_due))
    removed_summary = kpi_rules.removal_summary(
        removed_due, before_rows, len(report), len(eligible)
    )
    removed_detail = kpi_rules.removal_detail(removed)

    outside = pd.DataFrame()
    outside_path = DIRS["final_output"] / kpi_date / args.run_id / "status_check_outside_batch.csv"
    if outside_path.exists():
        outside = pd.read_csv(outside_path, dtype=str)

    exc_cols = ["TNO", "Source", "record_route", "进表来源", "kpi_start_time",
                "kpi_due_date", "final_ascan_time", "kpi_status", "status_exception"]
    exceptions = report[report["status_exception"].notna()][
        [c for c in exc_cols if c in report.columns]
    ] if "status_exception" in report.columns else report.iloc[0:0]

    ok, msg = pivots.check_dependencies()
    logger.info("Excel 依赖检查：%s", msg)
    if not ok:
        raise StopWithMessage(f"{msg}\n三个原生 Pivot 是硬性要求，不用静态汇总表冒充，因此中止。")

    out_dir = DIRS["final_output"] / kpi_date / args.run_id
    report_path = out_dir / "Ascan_KPI.xlsx"
    sheets = report_sheets(
        report, exceptions, outside, eligible, kpi_date, cfg,
        removed_summary, removed_detail,
    )
    report_path = pivots.build_workbook(report_path, sheets, cfg, logger)
    mia_path = None
    if bool(cfg.get("report", {}).get("write_mia_report", False)):
        mia_path = out_dir / pivots.MIA_REPORT_NAME
        mia_path = pivots.build_mia_hub_workbook(mia_path, report, cfg, logger)
        logger.info("MIA Hub 报表：%s", mia_path.relative_to(ROOT))
    report.to_parquet(step2, index=False)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch.update({
        "report_rebuilt_at": now,
        "report_rebuild_run_id": run_id2,
        "report": str(report_path.relative_to(ROOT)),
        "mia_report": str(mia_path.relative_to(ROOT)) if mia_path else None,
        "master_tno_total": int(report["TNO"].nunique()),
        "kpi_denominator": int(len(eligible)),
        "kpi_counts": {str(k): int(v) for k, v in counts.items()},
        "on_time_rate": rate,
    })
    write_json(work / "manifest.json", batch)
    write_json(out_dir / "manifest.json", batch)
    logger.info("报表已重出（未改 Pool）：%s  分母=%d  On Time Rate=%.4f%%",
                report_path.relative_to(ROOT), len(eligible), rate * 100)
    return 0


# ------------------------------------------------------------------------ main

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="第二步：更新状态并出报表")
    parser.add_argument("--run-id", default=None, help="第一步批次 run_id；有多个待处理批次时必填")
    parser.add_argument("--inspect-only", action="store_true",
                        help="只核对 Status Check 列名与覆盖率，不更新状态、不改 Pool、不归档")
    parser.add_argument("--allow-unconfirmed-receiving-date", action="store_true",
                        help="明知收货日期未确认仍要出报表（结果不得作为正式 KPI）")
    parser.add_argument("--rebuild-report", action="store_true",
                        help="已完成批次：只按当前规则重算 kpi_eligible 并重出报表，不改 Pool、不归档")
    parser.add_argument("--reuse-main-report", default=None,
                        help="已含 Pivot 的主报表路径；跳过 Ascan_KPI.xlsx 重建，继续 MIA / Pool / 归档")
    args = parser.parse_args(argv)

    ensure_dirs()
    run_id2 = new_run_id("step2")
    logger = setup_logging(run_id2, "step2")

    try:
        cfg = load_config("closed_loop.yaml")
        sc_cfg = load_config("status_check.yaml")

        if args.rebuild_report:
            return rebuild_completed_report(args, cfg, sc_cfg, logger, run_id2)

        batch = pick_batch(args.run_id, logger)
        run_id = batch["run_id"]
        kpi_date = batch["kpi_date"]
        work = DIRS["working"] / run_id
        logger.info("处理批次 run_id=%s（KPI 日期 %s），本次第二步 run=%s", run_id, kpi_date, run_id2)
        logger.info("主表含本次到期记录；Pivot 默认筛选 kpi_eligible=True。")

        date_ok, date_msg = receiving_date_ready(batch)
        logger.info("收货日期：%s", date_msg)
        if not args.inspect_only and not date_ok and not args.allow_unconfirmed_receiving_date:
            raise StopWithMessage(
                date_msg + "\n正式第二步在收货日期确认前不会运行。"
                "仅检查请加 --inspect-only。"
            )

        master = pd.read_parquet(work / "unified_master_step1.parquet")
        master, n_fbt_iah = kpi_rules.apply_hub_corrections(master)
        if n_fbt_iah:
            logger.info("TT FBT 无 DFW：Hub DFW -> IAH，共 %d 条", n_fbt_iah)
        master["kpi_eligible"] = kpi_rules.apply_eligibility_exclusions(
            kpi_rules.coerce_eligible(master["kpi_eligible"]), master
        )
        saved_removed = work / "removed_rows.parquet"
        if saved_removed.exists():
            removed = pd.read_parquet(saved_removed)
            before_rows = len(master) + int((removed.get("删除范围", pd.Series(dtype="object")) == "本次到期").sum()) if len(removed) else len(master)
            logger.info("第一步已删除 %d 条，第二步沿用该清单", len(removed))
        else:
            before_rows = len(master)
            master, removed = kpi_rules.split_removed_rows(master)
            if len(removed):
                removed["删除范围"] = "本次到期"
            logger.info("第一步主表删除 %d 条，剩余 %d", len(removed), len(master))
            for reason, n in (removed["删除原因"].value_counts().items() if len(removed) else []):
                logger.info("  删除 %s：%d", reason, int(n))
        removed_summary = kpi_rules.removal_summary(
            removed.loc[removed["删除范围"].eq("本次到期")] if len(removed) and "删除范围" in removed.columns else removed,
            before_rows,
            len(master),
            int(master["kpi_eligible"].fillna(False).astype(bool).sum()) if len(master) else 0,
        )
        removed_detail = kpi_rules.removal_detail(removed)
        logger.info("读取第一步主表 %d 行 / 唯一 TNO %d / kpi_eligible %d / Status_Check_List %d",
                    len(master), master["TNO"].nunique() if len(master) else 0,
                    int(master["kpi_eligible"].fillna(False).astype(bool).sum()) if len(master) else 0,
                    int(master["query_reason"].notna().sum()) if len(master) else 0)

        sc, sc_files, sc_stats = read_status_check(sc_cfg, logger)
        coverage = report_coverage(master, sc, logger)
        conflict_info = conflict_preview(master, sc, logger)
        inspect = {
            "run_id": run_id,
            "inspected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status_check_files": sc_stats["files"],
            "raw_rows": sc_stats["raw_rows"],
            "rows_with_tno": sc_stats["rows_with_tno"],
            "duplicate_tno_count": len(sc_stats["duplicate_tnos"]),
            "duplicate_tnos_sample": sc_stats["duplicate_tnos"][:20],
            "missing_logical_columns": sc_stats["missing_logical_columns"],
            "coverage": coverage,
            "ascan_conflicts": conflict_info,
            "ascan_conflict_policy": sc_cfg.get("ascan_conflict_policy", "keep_original"),
            "note": "预演合成数据产生的冲突不得当作真实业务冲突。此处统计来自本次放入 Input 的文件。",
        }
        write_json(work / "status_check_inspect.json", inspect)
        logger.info("覆盖率检查已写入 %s", (work / "status_check_inspect.json").relative_to(ROOT))

        if args.inspect_only:
            logger.info("inspect-only：到此停止。未更新状态，未改 Pool，未归档查询文件。")
            return 0

        missing_req = [k for k in REQUIRED_APPLY_FIELDS if k in sc_stats["missing_logical_columns"]]
        if missing_req:
            raise StopWithMessage(
                "Status Check 字段尚未核对完整，缺少："
                + "、".join(missing_req)
                + "。请先用 --inspect-only 核对真实列名，并在 config/status_check.yaml 补齐后再正式更新。"
            )

        if not date_ok and args.allow_unconfirmed_receiving_date:
            logger.warning("按要求带未确认收货日期继续，结果不得作为正式 KPI。")

        ok, msg = pivots.check_dependencies()
        logger.info("Excel 依赖检查：%s", msg)
        if not ok:
            raise StopWithMessage(
                f"{msg}\n三个原生 Pivot 是硬性要求，不用静态汇总表冒充，因此中止。"
            )

        updated = apply_status(master, sc, sc_cfg, logger)

        # Query rows that do not belong to this batch.
        known = set(master["TNO"].astype("string"))
        removed_ids = set(removed["TNO"].astype("string")) if len(removed) and "TNO" in removed.columns else set()
        outside = sc[~sc["TNO"].isin(known | removed_ids)].copy()
        if len(outside):
            logger.warning("查询结果里有 %d 个 TNO 不属于本批次，已写核对清单", len(outside))

        eligible, report, counts, rate = publish_kpi_frames(updated, sc_cfg, logger)
        if eligible.empty:
            raise StopWithMessage(
                f"本次 KPI 范围为空（kpi_date={kpi_date}）。请确认第一步的 --kpi-date 是否正确。"
            )
        updated = report

        exc_cols = ["TNO", "Source", "record_route", "进表来源", "kpi_start_time",
                    "kpi_due_date", "final_ascan_time", "kpi_status", "status_exception"]
        exceptions = updated[updated["status_exception"].notna()][
            [c for c in exc_cols if c in updated.columns]
        ]
        conflicts = updated[updated["ascan_conflict"].eq("Yes")][
            ["TNO", "original_ascan_time", "sc_ascan_time", "final_ascan_time",
             "ascan_source_final", "kpi_status"]
        ]

        out_dir = DIRS["final_output"] / kpi_date / run_id
        report_path = out_dir / "Ascan_KPI.xlsx"
        sheets = report_sheets(
            updated, exceptions, outside, eligible, kpi_date, cfg,
            removed_summary, removed_detail,
        )
        reuse = Path(args.reuse_main_report) if args.reuse_main_report else None
        if reuse is not None:
            if not reuse.exists():
                raise StopWithMessage(f"复用主报表不存在：{reuse}")
            report_path = pivots._copy_to_dest(reuse, report_path, logger)
            logger.info("复用已生成主报表：%s", report_path.relative_to(ROOT))
        else:
            report_path = pivots.build_workbook(report_path, sheets, cfg, logger)
            logger.info("报表已生成：%s", report_path.relative_to(ROOT))
        mia_path = None
        if bool(cfg.get("report", {}).get("write_mia_report", False)):
            mia_path = out_dir / pivots.MIA_REPORT_NAME
            mia_path = pivots.build_mia_hub_workbook(mia_path, updated, cfg, logger)
            logger.info("MIA Hub 报表：%s", mia_path.relative_to(ROOT))
        if not date_ok or args.allow_unconfirmed_receiving_date:
            logger.warning("非正式 KPI（收货日期未确认）。")

        # ---------------------------------------------------------- pool upkeep
        active, closed, _ = pool_mod.load_pool(logger)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        keep = eligible[still_tracking(eligible)].copy()
        rows = build_pool_rows(keep, kpi_date, run_id, now)
        logger.info("写回 Active_Pool %d 条（本次考核仍无 A-scan；未到期已在第一步入池）",
                    len(rows))
        active = pool_mod.upsert_active(active, rows, logger)

        done = eligible.loc[finished_tracking(eligible), "TNO"].astype("string")
        finished = active["TNO"].astype("string").isin(set(done))
        if int(finished.sum()):
            moving = active[finished].copy()
            moving["pool_state"] = pool_mod.STATE_CLOSED
            moving["closed_time"] = now
            moving["kpi_counted"] = "Yes"
            moving["counted_kpi_date"] = moving["counted_kpi_date"].fillna(kpi_date)
            moving["counted_run_id"] = moving["counted_run_id"].fillna(run_id)
            closed = pd.concat([closed, moving], ignore_index=True)
            active = active[~finished].copy()
            logger.info("已完成考核与追踪，移入 Closed_History %d 条", len(moving))
        pool_mod.save_pool(active, closed, logger, cfg, reason=f"step2 {run_id}")

        for name, frame in (
            ("status_exceptions.csv", exceptions),
            ("ascan_conflicts.csv", conflicts),
            ("status_check_outside_batch.csv", outside),
        ):
            atomic_write_csv(out_dir / name, frame)
        updated.to_parquet(work / "unified_master_step2.parquet", index=False)

        archived = []
        dest = DIRS["archive"] / "Status_Check" / run_id
        dest.mkdir(parents=True, exist_ok=True)
        for path in sc_files:
            shutil.move(str(path), str(dest / path.name))
            archived.append(path.name)
        logger.info("已归档本次查询结果 %d 个 -> %s", len(archived), dest.relative_to(ROOT))

        batch.update({
            "state": "completed",
            "step2_run_id": run_id2,
            "completed_at": now,
            "report": str(report_path.relative_to(ROOT)),
            "mia_report": str(mia_path.relative_to(ROOT)) if mia_path else None,
            "official_kpi": bool(date_ok) and not args.allow_unconfirmed_receiving_date,
            "master_tno_total": int(updated["TNO"].nunique()),
            "kpi_denominator": int(len(eligible)),
            "status_check_coverage": coverage,
            "ascan_conflicts": {
                "count": int(len(conflicts)),
                "policy": sc_cfg.get("ascan_conflict_policy", "keep_original"),
                "source": "real_status_check_input",
                "samples": conflict_info.get("conflict_samples", []),
            },
            "kpi_counts": {str(k): int(v) for k, v in counts.items()},
            "on_time_rate": rate,
            "archived_status_check": archived,
        })
        write_json(work / "manifest.json", batch)
        write_json(out_dir / "manifest.json", batch)

        logger.info("=" * 70)
        logger.info("第二步完成。报表：%s", report_path)
        logger.info("=" * 70)
        return 0
    except StopWithMessage as exc:
        logger.error("已停止：%s", exc)
        logger.error("本次未修改 Pool，也未归档任何输入，可修复后重跑。")
        return 1
    except Exception:
        logger.error("未处理错误：\n%s", traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
