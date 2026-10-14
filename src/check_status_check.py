# -*- coding: utf-8 -*-
"""Dry-run check of a real Status Check export. Reads only; changes nothing.

    python src/check_status_check.py [--run-id 20260908_001748_step1]

Coverage is always measured against the Status_Check_List TNO list for that batch,
never against the full master.
"""
from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

import stage2
from common import (
    DIRS,
    ROOT,
    StopWithMessage,
    atomic_write_csv,
    clean_text,
    ensure_dirs,
    find_header_row,
    load_config,
    new_run_id,
    norm_key,
    resolve_columns,
    setup_logging,
)


def raw_headers(cfg: dict, logger) -> list[tuple[str, list[str]]]:
    """Actual column names in each file, before any alias mapping."""
    out = []
    files = [p for p in sorted(DIRS["status_input"].iterdir())
             if p.is_file() and not p.name.startswith("~$")]
    scan = int(cfg.get("header_scan_rows", 6))
    for path in files:
        if path.suffix.lower() in {".xlsx", ".xlsm"}:
            wb = load_workbook(path, read_only=True, data_only=True)
            rows = [list(r) for r in wb[wb.sheetnames[0]].iter_rows(max_row=scan + 1,
                                                                    values_only=True)]
            wb.close()
        else:
            rows = None
            for enc in ("utf-8-sig", "utf-8", "gbk", "cp1252", "latin1"):
                try:
                    rows = pd.read_csv(path, dtype=str, header=None, encoding=enc,
                                       nrows=scan + 1).values.tolist()
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
        if not rows:
            raise StopWithMessage(f"查询结果文件为空：{path.name}")
        head = find_header_row(rows, cfg["columns"], scan)
        out.append((path.name, [clean_text(c) or "" for c in rows[head]]))
    return out


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Status Check 上线前核对（只读）")
    parser.add_argument("--run-id", default=None)
    args = parser.parse_args(argv)

    ensure_dirs()
    op_run = new_run_id("sccheck")
    logger = setup_logging(op_run, "sccheck")
    try:
        sc_cfg = load_config("status_check.yaml")
        batch = stage2.pick_batch(args.run_id, logger)
        run_id, kpi_date = batch["run_id"], batch["kpi_date"]
        work = DIRS["working"] / run_id
        logger.info("核对批次 %s（KPI 日期 %s）—— 只读，不写主表、不动 Pool、不归档",
                    run_id, kpi_date)

        # ---------------------------------------------------------- 列名核对
        logger.info("")
        logger.info("=== 1. 真实列名核对 ===")
        by_norm_alias = {}
        for logical, names in sc_cfg["columns"].items():
            for name in names:
                by_norm_alias[norm_key(name)] = logical
        for name, header in raw_headers(sc_cfg, logger):
            logger.info("  文件 %s，共 %d 列", name, len(header))
            matched, unmatched = [], []
            for col in header:
                if not col:
                    continue
                hit = by_norm_alias.get(norm_key(col))
                (matched if hit else unmatched).append(f"{col} -> {hit}" if hit else col)
            logger.info("    命中配置的列：%s", matched or "（无）")
            logger.info("    未使用的列  ：%s", unmatched or "（无）")

        sc, sc_files, stats = stage2.read_status_check(sc_cfg, logger)
        if stats["missing_logical_columns"]:
            logger.warning("  配置里没有匹配上的字段：%s", stats["missing_logical_columns"])
        for logical, col in (("A-scan 时间", "sc_ascan_time"),
                             ("事件码", "sc_event_code"),
                             ("事件描述", "sc_event_desc"),
                             ("Last Status Event Description", "sc_status_event_desc"),
                             ('Last "Attempt" Event Description', "sc_attempt_event_desc"),
                             ("Last Public Status Event Description", "sc_public_status_event_desc")):
            if sc[col].isna().all():
                logger.warning("  %s 整列为空，相关判定会退化", logical)

        # ---------------------------------------------------------- 覆盖率
        master = pd.read_parquet(work / "unified_master_step1.parquet")
        master["TNO"] = master["TNO"].astype("string")
        query = master[master["query_reason"].notna()].drop_duplicates("TNO")
        query_ids = set(query["TNO"])
        master_ids = set(master["TNO"])
        sc_ids = set(sc["TNO"].astype("string"))

        hit = query_ids & sc_ids
        miss = query_ids - sc_ids
        extra = sc_ids - query_ids
        extra_in_master = extra & master_ids
        extra_unknown = extra - master_ids

        denom = max(len(query_ids), 1)
        logger.info("")
        logger.info("=== 2. 匹配覆盖率（分母 = Status_Check_List 唯一 TNO）===")
        logger.info("  完整主表 TNO 数        ： %8d  （仅作对照，不是分母）", len(master_ids))
        logger.info("  本次 KPI 分母          ： %8d  （仅作对照，不是分母）",
                    int(master["kpi_eligible"].fillna(False).astype(bool).sum()))
        logger.info("  ------------------------------------------")
        logger.info("  应查询数量             ： %8d", len(query_ids))
        logger.info("  匹配成功数量           ： %8d  (%.2f%%)", len(hit), 100 * len(hit) / denom)
        logger.info("  缺失数量               ： %8d  (%.2f%%)", len(miss), 100 * len(miss) / denom)
        logger.info("  多余数量               ： %8d", len(extra))
        logger.info("    其中在主表但不在清单 ： %8d", len(extra_in_master))
        logger.info("    其中完全不在主表     ： %8d", len(extra_unknown))

        for reason, part in query.groupby("query_reason"):
            ids = set(part["TNO"])
            logger.info("    [%s] 应查 %d，命中 %d，缺失 %d",
                        reason, len(ids), len(ids & sc_ids), len(ids - sc_ids))

        # ---------------------------------------------------------- 重复与冲突
        dup_ids = stats["duplicate_tnos"]
        joined = query.merge(sc, on="TNO", how="inner")
        orig = pd.to_datetime(joined["first_ascan_time"], errors="coerce")
        got = pd.to_datetime(joined["sc_ascan_time"], errors="coerce")
        conflict = joined[orig.notna() & got.notna() & (orig != got)].copy()

        logger.info("")
        logger.info("=== 3. 重复与冲突 ===")
        logger.info("  查询结果原始行数       ： %8d", stats["raw_rows"])
        logger.info("  查询结果内重复 TNO     ： %8d", len(dup_ids))
        logger.info("  原A-scan与查询A-scan冲突： %8d", len(conflict))
        logger.info("  冲突处理策略           ： %s（不自行改时间优先级）",
                    sc_cfg.get("ascan_conflict_policy"))
        if len(conflict):
            sample = conflict[["TNO", "first_ascan_time", "sc_ascan_time",
                               "kpi_due_date", "status"]].head(10)
            logger.info("  冲突样例（原时间 vs 查询时间）：")
            for _, r in sample.iterrows():
                logger.info("    %s  原=%s  查询=%s  截止=%s  原状态=%s",
                            r["TNO"], r["first_ascan_time"], r["sc_ascan_time"],
                            r["kpi_due_date"], r["status"])

        # ---------------------------------------------------------- 输出
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = work / f"status_check_precheck_{stamp}"
        atomic_write_csv(out.with_name(out.name + "_missing.csv"),
                         query[query["TNO"].isin(miss)][
                             ["TNO", "query_reason", "Source", "Hub", "CUSTOMER",
                              "kpi_due_date", "status"]])
        atomic_write_csv(out.with_name(out.name + "_extra.csv"),
                         sc[sc["TNO"].isin(extra)])
        atomic_write_csv(out.with_name(out.name + "_conflicts.csv"),
                         conflict[["TNO", "first_ascan_time", "sc_ascan_time",
                                   "kpi_start_time", "kpi_due_date", "status",
                                   "sc_event_code", "sc_event_desc"]])
        logger.info("")
        logger.info("明细已写入 %s_*.csv（只读核对，未改动任何正式数据）",
                    out.relative_to(ROOT))
        if miss or extra_unknown or dup_ids:
            logger.warning("存在缺失/多余/重复，请先核对再跑正式第二步。")
        return 0
    except StopWithMessage as exc:
        logger.error("已停止：%s", exc)
        return 1
    except Exception:
        logger.error("未处理错误：\n%s", traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
