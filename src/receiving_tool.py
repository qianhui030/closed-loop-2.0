# -*- coding: utf-8 -*-
"""Confirm or correct receiving dates in place. Never create a second dated copy."""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

import receiving_log as rl
from common import DIRS, StopWithMessage, ensure_dirs, load_config, setup_logging


def _logger() -> logging.Logger:
    ensure_dirs()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return setup_logging(stamp, "receiving_tool")


def cmd_confirm(args) -> int:
    logger = _logger()
    rl.revise_receiving_date(args.source_file, args.date, "confirm", logger, note=args.note or "用户确认")
    logger.info("该文件收货日期已确认。正式第二步现在可以继续。")
    return 0


def cmd_correct(args) -> int:
    logger = _logger()
    rl.revise_receiving_date(args.source_file, args.new_date, "correct", logger, note=args.note or "用户更正")
    logger.info("已按原事件就地改日期，没有新增重复收货事件。修改记录在 Receiving_Log/Standardized/receiving_date_revisions.csv")
    return 0


def cmd_backfill(args) -> int:
    """Rebuild missing columns / no-MAWB rows for an already archived receiving file."""
    logger = _logger()
    cfg = load_config("receiving_log.yaml")
    path = Path(args.file)
    if not path.exists():
        # search archive
        hits = list(DIRS["archive"].rglob(args.file))
        if not hits:
            raise StopWithMessage(f"找不到收货文件 {args.file}")
        path = hits[0]
    date = datetime.strptime(args.receiving_date, "%Y-%m-%d").date()
    std = rl.standardize_file(path, cfg, args.batch_id, date, logger)
    history = rl.hydrate_existing(std, logger)
    pending = list(std.loc[std["receiving_date_status"] != rl.DATE_CONFIRMED, "source_file"].unique())
    rl._mark_manifests_date(path.name, confirmed=False, argument=args.receiving_date)

    work = DIRS["working"] / args.batch_id
    if work.exists():
        held = history[history["match_eligible"] != "Yes"]
        from common import atomic_write_csv
        atomic_write_csv(work / "receiving_not_matchable.csv", held)
        multi = (
            history[history["mawb_key"].notna()]
            .groupby("mawb_key")
            .filter(lambda g: g["receiving_record_id"].nunique() > 1)
            [["MAWB", "mawb_key", "mawb_format", "receiving_time", "source_row",
              "match_status", "match_eligible", "receiving_record_id"]]
            .sort_values(["mawb_key", "receiving_time"])
        )
        atomic_write_csv(work / "receiving_split_arrivals.csv", multi)

    logger.info("回填完成。历史 %d 条。日期仍为待确认：%s", len(history), pending)
    logger.info("确认：python src\\receiving_tool.py confirm --source-file \"%s\"", path.name)
    logger.info("更正：python src\\receiving_tool.py correct --source-file \"%s\" --new-date YYYY-MM-DD", path.name)
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="收货日期确认/更正（就地修改，不重复导入）")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_cf = sub.add_parser("confirm", help="确认该文件当前收货日期，不再标记待确认")
    p_cf.add_argument("--source-file", required=True)
    p_cf.add_argument("--date", default=None, help="YYYY-MM-DD；省略则沿用历史中的唯一日期")
    p_cf.add_argument("--note", default="")
    p_cf.set_defaults(func=cmd_confirm)

    p_co = sub.add_parser("correct", help="更正该文件收货日期，保留原事件与修改记录")
    p_co.add_argument("--source-file", required=True)
    p_co.add_argument("--new-date", required=True, help="YYYY-MM-DD")
    p_co.add_argument("--note", default="")
    p_co.set_defaults(func=cmd_correct)

    p_bf = sub.add_parser("backfill", help="从归档文件补全历史字段与无MAWB行")
    p_bf.add_argument("--file", required=True)
    p_bf.add_argument("--batch-id", required=True)
    p_bf.add_argument("--receiving-date", required=True)
    p_bf.set_defaults(func=cmd_backfill)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except StopWithMessage as exc:
        print(f"已停止：{exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
