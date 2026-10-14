# -*- coding: utf-8 -*-
"""Receiving Log standardization plus a cumulative history that survives archiving."""
from __future__ import annotations

import fnmatch
import hashlib
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from common import (
    DIRS,
    RECEIVING_HISTORY,
    StopWithMessage,
    atomic_write_csv,
    backup_file,
    clean_text,
    find_header_row,
    mawb_match_key,
    parse_dt,
    resolve_columns,
)

STD_COLS = [
    "receiving_record_id", "event_key", "content_key", "MAWB", "mawb_raw", "mawb_key", "mawb_format",
    "Hub", "hub_source", "receiving_time", "receiving_date_source",
    "receiving_date_status", "cartons", "skids",
    "broker", "first_mile_type", "receiver",
    "source_file", "source_row", "batch_id", "imported_at",
    "match_status", "match_eligible", "field_exception",
    "date_revision", "revision_note",
]
REVISION_LOG = DIRS["receiving_std"] / "receiving_date_revisions.csv"
REVISION_COLS = [
    "revision_id", "revised_at", "action", "source_file", "old_date", "new_date",
    "receiving_record_id", "event_key", "MAWB", "old_receiving_time",
    "new_receiving_time", "note",
]

MAWB_FORMAT_AWB = "AWB"
MAWB_FORMAT_UNKNOWN = "待确认_非标准单号"
DATE_CONFIRMED = "已确认"
DATE_PENDING = "待确认"
DATE_UNKNOWN = "无法确定"

# Air waybill: three-digit prefix plus eight digits, dash optional.
AWB_PATTERN = re.compile(r"^\d{3}-?\d{8}$")


def classify_mawb(value) -> str | None:
    """Never rewrite the number. Just label anything that is not a plain AWB."""
    text = clean_text(value)
    if text is None:
        return None
    return MAWB_FORMAT_AWB if AWB_PATTERN.fullmatch(text) else MAWB_FORMAT_UNKNOWN


def date_status_for(source: str | None) -> str:
    if source in {"file_datetime", "file_date"}:
        return DATE_CONFIRMED
    if source == "run_argument":
        return DATE_PENDING
    return DATE_UNKNOWN


def _read_table(path: Path, cfg: dict) -> pd.DataFrame:
    """Read csv/xlsx as raw text, then locate the real header row."""
    aliases = cfg["columns"]
    scan = int(cfg.get("header_scan_rows", 6))
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    else:
        last_error = None
        rows = None
        for enc in cfg.get("encodings", ["utf-8-sig"]):
            try:
                raw = pd.read_csv(path, dtype=str, header=None, encoding=enc,
                                  keep_default_na=False, na_values=[""])
                rows = raw.values.tolist()
                break
            except (UnicodeDecodeError, UnicodeError) as exc:
                last_error = exc
        if rows is None:
            raise StopWithMessage(f"无法解码收货文件 {path.name}：{last_error}")
    if not rows:
        raise StopWithMessage(f"收货文件为空：{path.name}")

    header_idx = find_header_row(rows, aliases, scan)
    header = rows[header_idx]
    # A two-row header keeps MAWB on the Chinese line and the rest on the English line.
    merged = []
    for col_idx, cell in enumerate(header):
        name = clean_text(cell)
        if name is None:
            for other in range(min(scan, len(rows))):
                if other == header_idx:
                    continue
                alt = clean_text(rows[other][col_idx]) if col_idx < len(rows[other]) else None
                if alt is not None:
                    name = alt
                    break
        merged.append(name if name is not None else f"col_{col_idx}")
    body = rows[header_idx + 1:]
    df = pd.DataFrame(body, columns=merged)
    df["__source_row"] = range(header_idx + 2, header_idx + 2 + len(df))
    df = df.dropna(how="all", subset=[c for c in merged])
    return df


def _resolve_hub(path: Path, cfg: dict, in_row_hub) -> tuple[str | None, str]:
    hub = clean_text(in_row_hub)
    if hub:
        return hub.upper(), "file_column"
    for rule in cfg.get("file_hub_map", []) or []:
        if fnmatch.fnmatch(path.name, rule.get("pattern", "")):
            return str(rule["hub"]).upper(), "config:file_hub_map"
    return None, "unresolved"


def _parse_receiving_time(row, cols, cfg, external_date) -> tuple[pd.Timestamp, str, str | None]:
    """Full datetime > date column + time column > external date + time column."""
    if "receiving_datetime" in cols:
        value = parse_dt(row.get(cols["receiving_datetime"]), cfg.get("datetime_formats"))
        if pd.notna(value):
            return value, "file_datetime", None

    time_text = clean_text(row.get(cols["receiving_time_only"])) if "receiving_time_only" in cols else None
    time_part = None
    if time_text:
        for fmt in cfg.get("time_formats", []):
            try:
                time_part = datetime.strptime(time_text, fmt).time()
                break
            except ValueError:
                continue
        if time_part is None:
            return pd.NaT, "unparsed", f"时间无法解析：{time_text!r}"

    date_part = None
    date_source = None
    if "receiving_date" in cols:
        raw_date = clean_text(row.get(cols["receiving_date"]))
        if raw_date:
            for fmt in cfg.get("date_formats", []):
                try:
                    date_part = datetime.strptime(raw_date, fmt).date()
                    date_source = "file_date"
                    break
                except ValueError:
                    continue
            if date_part is None:
                return pd.NaT, "unparsed", f"日期无法解析：{raw_date!r}"
    if date_part is None and external_date is not None:
        date_part = external_date
        date_source = "run_argument"

    if time_part is None:
        return pd.NaT, "missing_time", "文件内没有可用的收货时间"
    if date_part is None:
        return pd.NaT, "missing_date", (
            "文件内没有收货日期列，且未提供 --receiving-date。"
            "按规则不使用文件修改时间，也不从文件名推断。"
        )
    return pd.Timestamp(datetime.combine(date_part, time_part)), date_source, None


def _to_int(value) -> tuple[int | None, str | None]:
    text = clean_text(value)
    if text is None:
        return None, None
    if re.fullmatch(r"\d+(\.0+)?", text):
        return int(float(text)), None
    return None, f"数量字段不是整数：{text!r}"


def _record_id(hub, mawb_key, receiving_time, cartons, skids, broker,
               source_file=None, source_row=None) -> str:
    """Content hash so re-importing the same log does not create a second event.

    Rows without a MAWB have no other discriminator, so they fall back to the
    file and row number. That is still stable across re-imports of the same file.
    """
    parts = [
        str(hub or ""),
        str(mawb_key or ""),
        receiving_time.strftime("%Y-%m-%d %H:%M:%S") if pd.notna(receiving_time) else "",
        str(cartons if cartons is not None else ""),
        str(skids if skids is not None else ""),
        str(broker or ""),
    ]
    if not mawb_key:
        parts += [str(source_file or ""), str(source_row or "")]
    return "RCV" + hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:16]


def event_key(source_file, source_row, mawb_key) -> str:
    """Identity for one exploded cell in one source file. Used when correcting dates."""
    return f"{source_file or ''}|{source_row or ''}|{mawb_key or ''}"


def content_key(hub, mawb_key, receiving_time, cartons, skids, broker, source_row=None) -> str:
    """Date-independent fingerprint of a physical arrival.

    Does not include the file name, so renaming or re-importing the same log
    cannot create a second event. The calendar date is excluded so a later
    date correction does not look like a new arrival.
    """
    tod = ""
    if pd.notna(receiving_time):
        tod = pd.Timestamp(receiving_time).strftime("%H:%M:%S")
    base = [
        str(hub or "").upper(),
        str(mawb_key or ""),
        tod,
        str(cartons if cartons is not None and not (isinstance(cartons, float) and pd.isna(cartons)) else ""),
        str(skids if skids is not None and not (isinstance(skids, float) and pd.isna(skids)) else ""),
        str(broker or ""),
    ]
    if not mawb_key:
        base.append(str(source_row or ""))
    return "|".join(base)


def standardize_file(path: Path, cfg: dict, batch_id: str, external_date, logger) -> pd.DataFrame:
    df = _read_table(path, cfg)
    cols = resolve_columns([c for c in df.columns if c != "__source_row"], cfg["columns"])
    logger.info("  %s -> 命中列 %s", path.name, {k: v for k, v in cols.items()})
    if "mawb" not in cols:
        raise StopWithMessage(
            f"收货文件 {path.name} 找不到 MAWB 列。已识别的列：{list(df.columns)}\n"
            "请在 config/receiving_log.yaml 的 columns.mawb 里补上真实列名。"
        )
    split_pat = cfg.get("mawb_split_pattern", r"[\r\n;,]+")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    out = []
    for _, row in df.iterrows():
        # Blocking problems stop MAWB matching. Data-quality notes are recorded only.
        blocking: list[str] = []
        notes: list[str] = []
        hub, hub_source = _resolve_hub(path, cfg, row.get(cols["hub"]) if "hub" in cols else None)
        if hub is None:
            blocking.append("Hub 未确定：表内无 Hub 列且未命中 file_hub_map 配置")
        rec_time, date_source, time_err = _parse_receiving_time(row, cols, cfg, external_date)
        if time_err:
            blocking.append(time_err)
        cartons, carton_err = _to_int(row.get(cols["cartons"])) if "cartons" in cols else (None, None)
        skids, skid_err = _to_int(row.get(cols["skids"])) if "skids" in cols else (None, None)
        for err in (carton_err, skid_err):
            if err:
                notes.append(err)
        broker = clean_text(row.get(cols["broker"])) if "broker" in cols else None
        fmt_type = clean_text(row.get(cols["first_mile_type"])) if "first_mile_type" in cols else None
        receiver = clean_text(row.get(cols["receiver"])) if "receiver" in cols else None

        date_status = date_status_for(date_source)
        base = {
            "Hub": hub, "hub_source": hub_source,
            "receiving_time": rec_time, "receiving_date_source": date_source,
            "receiving_date_status": date_status,
            "cartons": cartons, "skids": skids,
            "broker": broker, "first_mile_type": fmt_type, "receiver": receiver,
            "source_file": path.name, "source_row": int(row["__source_row"]),
            "batch_id": batch_id, "imported_at": now,
            "date_revision": 0, "revision_note": None,
        }

        raw_mawb = clean_text(row.get(cols["mawb"]))
        pieces = [p.strip() for p in re.split(split_pat, raw_mawb)] if raw_mawb else []
        pieces = [p for p in pieces if p]
        if not pieces:
            # Kept in the history on purpose: it is a real receiving event that
            # simply cannot be matched to packages yet.
            out.append({
                **base,
                "receiving_record_id": _record_id(hub, None, rec_time, cartons, skids,
                                                  broker, path.name, row["__source_row"]),
                "event_key": event_key(path.name, row["__source_row"], None),
                "content_key": content_key(hub, None, rec_time, cartons, skids, broker, row["__source_row"]),
                "MAWB": None, "mawb_raw": raw_mawb, "mawb_key": None, "mawb_format": None,
                "match_status": "待核对_无MAWB", "match_eligible": "No",
                "field_exception": ";".join(
                    blocking + notes + ["该行没有 MAWB，无法匹配到包裹"]) or None,
            })
            continue
        for piece in pieces:
            key = mawb_match_key(piece)
            fmt = classify_mawb(piece)
            row_block = list(blocking)
            if fmt == MAWB_FORMAT_UNKNOWN:
                row_block.append("非航空主单格式，单号原值保留，与包裹的关联方式待确认")
                status = "待核对_非标准单号"
            elif row_block:
                status = "待核对"
            else:
                status = "已标准化"
            eligible = "No" if row_block else "Yes"
            out.append({
                **base,
                "receiving_record_id": _record_id(hub, key, rec_time, cartons, skids, broker),
                "event_key": event_key(path.name, row["__source_row"], key),
                "content_key": content_key(hub, key, rec_time, cartons, skids, broker),
                "MAWB": piece, "mawb_raw": raw_mawb, "mawb_key": key, "mawb_format": fmt,
                "match_status": status, "match_eligible": eligible,
                "field_exception": ";".join(row_block + notes) or None,
            })
    return pd.DataFrame(out, columns=STD_COLS)


def standardize_inputs(cfg: dict, batch_id: str, external_date, logger) -> tuple[pd.DataFrame, list[Path]]:
    files = [p for p in sorted(DIRS["receiving_input"].iterdir())
             if p.is_file() and not p.name.startswith("~$")] if DIRS["receiving_input"].exists() else []
    if not files:
        logger.warning("Receiving_Log/Input 为空，本次不写入任何收货时间。")
        return pd.DataFrame(columns=STD_COLS), []
    logger.info("标准化收货文件 %d 个", len(files))
    parts = [standardize_file(p, cfg, batch_id, external_date, logger) for p in files]
    std = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=STD_COLS)

    dup = std["receiving_record_id"].duplicated()
    if dup.any():
        logger.warning("同一批次内有 %d 条完全相同的收货记录，已合并为一条事件", int(dup.sum()))
        std = std[~dup].copy()

    batch_path = DIRS["receiving_std"] / f"receiving_standardized_{batch_id}.csv"
    atomic_write_csv(batch_path, std)
    logger.info("批次标准化结果：%s（%d 条）", batch_path.name, len(std))
    for label, count in std["match_status"].value_counts().items():
        logger.info("    %s: %d", label, count)
    pending_date = int((std["receiving_date_status"] != DATE_CONFIRMED).sum())
    if pending_date:
        logger.warning("有 %d 条收货记录的日期来自运行参数，标记为「%s」。"
                       "确认或更正前不要出正式 KPI。", pending_date, DATE_PENDING)
    return std, files


def load_history() -> pd.DataFrame:
    if not RECEIVING_HISTORY.exists():
        return pd.DataFrame(columns=STD_COLS)
    history = pd.read_csv(RECEIVING_HISTORY, dtype=str)
    for col in STD_COLS:
        if col not in history.columns:
            history[col] = pd.NA
    history["receiving_time"] = pd.to_datetime(history["receiving_time"], errors="coerce")
    missing_key = history["event_key"].isna() | (history["event_key"].astype("string").str.len() == 0)
    if missing_key.any():
        history.loc[missing_key, "event_key"] = [
            event_key(f, r, k) for f, r, k in zip(
                history.loc[missing_key, "source_file"],
                history.loc[missing_key, "source_row"],
                history.loc[missing_key, "mawb_key"],
            )
        ]
    missing_ck = history["content_key"].isna() | (history["content_key"].astype("string").str.len() == 0)
    if missing_ck.any():
        history.loc[missing_ck, "content_key"] = [
            content_key(h, k, t, c, s, b, r) for h, k, t, c, s, b, r in zip(
                history.loc[missing_ck, "Hub"],
                history.loc[missing_ck, "mawb_key"],
                history.loc[missing_ck, "receiving_time"],
                history.loc[missing_ck, "cartons"],
                history.loc[missing_ck, "skids"],
                history.loc[missing_ck, "broker"],
                history.loc[missing_ck, "source_row"],
            )
        ]
    return history[STD_COLS + [c for c in history.columns if c not in STD_COLS]]


def append_history(std: pd.DataFrame, logger) -> pd.DataFrame:
    """Cumulative store. Every event is kept, including rows that cannot be matched.

    Dedup is by event_key (file + row + MAWB piece), not by the dated hash.
    Re-importing the same log with a different --receiving-date therefore cannot
    create a second event; use revise_receiving_date() to change the date.
    """
    history = load_history()
    if std.empty:
        return history
    incoming = std.copy()
    if "event_key" not in incoming.columns or incoming["event_key"].isna().all():
        incoming["event_key"] = [
            event_key(f, r, k) for f, r, k in zip(
                incoming.get("source_file"), incoming.get("source_row"), incoming.get("mawb_key")
            )
        ]
    if "content_key" not in incoming.columns or incoming["content_key"].isna().all():
        incoming["content_key"] = [
            content_key(h, k, t, c, s, b, r) for h, k, t, c, s, b, r in zip(
                incoming.get("Hub"), incoming.get("mawb_key"), incoming.get("receiving_time"),
                incoming.get("cartons"), incoming.get("skids"), incoming.get("broker"),
                incoming.get("source_row"),
            )
        ]
    known_ids = set(history["receiving_record_id"].dropna().astype(str)) if len(history) else set()
    known_keys = set(history["event_key"].dropna().astype(str)) if len(history) and "event_key" in history.columns else set()
    known_content = set(history["content_key"].dropna().astype(str)) if len(history) and "content_key" in history.columns else set()

    by_id = incoming["receiving_record_id"].astype("string").isin(known_ids)
    by_key = incoming["event_key"].astype("string").isin(known_keys)
    by_content = incoming["content_key"].astype("string").isin(known_content) if "content_key" in incoming.columns else False
    skip = by_id | by_key | by_content
    if skip.any() and len(history):
        skipped = incoming[skip]
        hist_times = history.set_index("event_key")["receiving_time"] if "event_key" in history.columns else pd.Series(dtype="datetime64[ns]")
        for _, row in skipped.iterrows():
            old = hist_times.get(row["event_key"]) if len(hist_times) else None
            new = row.get("receiving_time")
            if pd.notna(old) and pd.notna(new) and pd.Timestamp(old).normalize() != pd.Timestamp(new).normalize():
                logger.warning(
                    "收货事件已存在但日期不同，未新增重复：%s %s  历史=%s  本次=%s。"
                    "若要改日期，请运行 python src\\receiving_tool.py correct --source-file \"%s\" --new-date YYYY-MM-DD",
                    row.get("source_file"), row.get("MAWB") or "(无MAWB)",
                    pd.Timestamp(old).strftime("%Y-%m-%d"),
                    pd.Timestamp(new).strftime("%Y-%m-%d"),
                    row.get("source_file"),
                )
    fresh = incoming[~skip]
    merged = pd.concat([history, fresh], ignore_index=True) if len(fresh) else history
    atomic_write_csv(RECEIVING_HISTORY, merged)
    logger.info("收货历史累计 %d 条（本次新增 %d，重复导入忽略 %d）",
                len(merged), len(fresh), int(skip.sum()))
    if "match_eligible" in merged.columns:
        logger.info("  其中可自动匹配 Pool 的事件 %d 条，待核对保留 %d 条",
                    int((merged["match_eligible"] == "Yes").sum()),
                    int((merged["match_eligible"] != "Yes").sum()))
    return merged


def _append_revision_log(rows: list[dict]) -> None:
    existing = pd.read_csv(REVISION_LOG, dtype=str) if REVISION_LOG.exists() else pd.DataFrame(columns=REVISION_COLS)
    atomic_write_csv(REVISION_LOG, pd.concat([existing, pd.DataFrame(rows, columns=REVISION_COLS)], ignore_index=True))


def _select_by_source(history: pd.DataFrame, source_file: str) -> pd.Series:
    names = history["source_file"].astype("string")
    hit = names.eq(source_file) | names.str.endswith(source_file, na=False)
    if not hit.any():
        available = sorted(set(names.dropna()))
        raise StopWithMessage(
            f"收货历史中找不到来源文件 {source_file!r}。已有文件：{available}"
        )
    return hit


def revise_receiving_date(source_file: str, new_date, action: str, logger, note: str = "") -> pd.DataFrame:
    """Confirm or correct the calendar date of already-stored events. In-place only."""
    if action not in {"confirm", "correct"}:
        raise StopWithMessage(f"未知操作 {action}")
    if new_date is not None and isinstance(new_date, str):
        new_date = datetime.strptime(new_date, "%Y-%m-%d").date()
    if action == "correct" and new_date is None:
        raise StopWithMessage("更正日期必须提供 --new-date YYYY-MM-DD")
    history = load_history()
    if history.empty:
        raise StopWithMessage("收货历史为空，没有可更正的事件。")
    history = history.astype(object)
    mask = _select_by_source(history, source_file)
    subset = history.loc[mask].copy()
    old_dates = pd.to_datetime(subset["receiving_time"], errors="coerce").dt.strftime("%Y-%m-%d")
    if action == "confirm" and new_date is None:
        unique_dates = sorted(set(old_dates.dropna()))
        if len(unique_dates) != 1:
            raise StopWithMessage(
                f"该文件的收货日期不唯一 {unique_dates}，确认时请显式给出 --date YYYY-MM-DD。"
            )
        new_date = datetime.strptime(unique_dates[0], "%Y-%m-%d").date()
    if action == "confirm":
        other = sorted({d for d in old_dates.dropna() if d != new_date.isoformat()})
        if other:
            raise StopWithMessage(
                f"历史日期 {sorted(set(old_dates.dropna()))} 与确认日期 {new_date} 不一致。"
                "若要改日期，请用 correct，不要用 confirm。"
            )
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file(RECEIVING_HISTORY, DIRS["receiving_std"] / "History_Backups", keep=30)

    log_rows = []
    for idx, row in subset.iterrows():
        old_ts = pd.to_datetime(row["receiving_time"], errors="coerce")
        if pd.isna(old_ts):
            continue
        new_ts = pd.Timestamp(datetime.combine(new_date, old_ts.time()))
        if action == "correct":
            history.at[idx, "receiving_time"] = new_ts
        history.at[idx, "receiving_date_status"] = DATE_CONFIRMED
        history.at[idx, "receiving_date_source"] = (
            "user_confirmed" if action == "confirm" else "user_corrected"
        )
        prev = int(pd.to_numeric(row.get("date_revision"), errors="coerce") or 0)
        bumped = 0 if action == "confirm" and old_ts.date() == new_date else 1
        history.at[idx, "date_revision"] = str(prev + bumped)
        history.at[idx, "revision_note"] = (
            f"{action} {old_ts.strftime('%Y-%m-%d')} -> {new_date.isoformat()} @ {now}"
            + (f"；{note}" if note else "")
        )
        log_rows.append({
            "revision_id": f"{stamp}_{row['receiving_record_id']}",
            "revised_at": now, "action": action, "source_file": row.get("source_file"),
            "old_date": old_ts.strftime("%Y-%m-%d"), "new_date": new_date.isoformat(),
            "receiving_record_id": row.get("receiving_record_id"),
            "event_key": row.get("event_key"), "MAWB": row.get("MAWB"),
            "old_receiving_time": old_ts.strftime("%Y-%m-%d %H:%M:%S"),
            "new_receiving_time": new_ts.strftime("%Y-%m-%d %H:%M:%S") if action == "correct" else old_ts.strftime("%Y-%m-%d %H:%M:%S"),
            "note": note or "",
        })

    atomic_write_csv(RECEIVING_HISTORY, history)
    _append_revision_log(log_rows)
    _sync_batch_copies(history)
    times_changed = any(
        pd.Timestamp(r["old_receiving_time"]) != pd.Timestamp(r["new_receiving_time"])
        for r in log_rows
    )
    if times_changed:
        _sync_pool_receiving_times(history.loc[mask], logger)
    else:
        logger.info("收货时间未改，跳过 Pool 整表读写")
    _mark_manifests_date(source_file, confirmed=True, argument=new_date.isoformat())

    logger.info("%s 收货日期：文件 %s，事件 %d 条，日期 %s -> %s。修改记录：%s",
                "确认" if action == "confirm" else "更正",
                source_file, int(mask.sum()),
                ", ".join(sorted(set(old_dates.dropna()))),
                new_date.isoformat(), REVISION_LOG.name)
    return history


def hydrate_existing(std: pd.DataFrame, logger) -> pd.DataFrame:
    """Fill new columns onto already-stored events and append any missing rows.

    Used to backfill the current batch without creating a second dated copy.
    Existing receiving_record_id values are kept.
    """
    history = load_history()
    if std.empty:
        return history
    incoming = std.copy()
    incoming["event_key"] = [
        event_key(f, r, k) for f, r, k in zip(
            incoming["source_file"], incoming["source_row"], incoming["mawb_key"]
        )
    ]
    incoming["content_key"] = [
        content_key(h, k, t, c, s, b, r) for h, k, t, c, s, b, r in zip(
            incoming["Hub"], incoming["mawb_key"], incoming["receiving_time"],
            incoming["cartons"], incoming["skids"], incoming["broker"], incoming["source_row"],
        )
    ]
    if history.empty:
        atomic_write_csv(RECEIVING_HISTORY, incoming)
        return incoming
    history = history.copy().astype(object)
    incoming = incoming.astype(object)
    history["event_key"] = history["event_key"].fillna(
        pd.Series([
            event_key(f, r, k) for f, r, k in zip(
                history["source_file"], history["source_row"], history["mawb_key"]
            )
        ], index=history.index)
    )
    by_key = incoming.set_index("event_key", drop=False)
    overlay = [
        "mawb_format", "receiving_date_status", "match_status", "match_eligible",
        "field_exception", "event_key", "content_key", "MAWB", "mawb_raw", "mawb_key",
        "Hub", "hub_source", "receiving_time", "receiving_date_source",
        "cartons", "skids", "broker", "first_mile_type", "receiver",
        "source_file", "source_row", "batch_id",
    ]
    matched = 0
    for idx, row in history.iterrows():
        key = row.get("event_key")
        if key not in by_key.index:
            continue
        src = by_key.loc[key]
        if isinstance(src, pd.DataFrame):
            src = src.iloc[0]
        for col in overlay:
            if col in src.index and pd.notna(src[col]):
                history.at[idx, col] = src[col]
        if pd.isna(row.get("date_revision")):
            history.at[idx, "date_revision"] = 0
        matched += 1
    known = set(history["event_key"].dropna().astype(str))
    fresh = incoming[~incoming["event_key"].astype("string").isin(known)]
    merged = pd.concat([history, fresh], ignore_index=True) if len(fresh) else history
    atomic_write_csv(RECEIVING_HISTORY, merged)
    logger.info("回填收货历史：更新已有 %d 条，补入缺失 %d 条，合计 %d",
                matched, len(fresh), len(merged))
    _sync_batch_copies(merged)
    return merged


def _sync_batch_copies(history: pd.DataFrame) -> None:
    """Keep per-batch standardized files aligned after an in-place date edit."""
    if "batch_id" not in history.columns:
        return
    for batch_id, part in history.groupby(history["batch_id"].astype("string")):
        if not batch_id or batch_id == "<NA>":
            continue
        path = DIRS["receiving_std"] / f"receiving_standardized_{batch_id}.csv"
        if path.exists():
            atomic_write_csv(path, part[STD_COLS] if set(STD_COLS) <= set(part.columns) else part)


def _sync_pool_receiving_times(changed: pd.DataFrame, logger) -> None:
    """If Pool already copied a receiving_time, keep it in lockstep. Do not recount KPI."""
    from pool import pool_store_exists, load_pool, save_pool  # local import to avoid a cycle at module load

    if changed.empty or not pool_store_exists():
        return
    ids = set(changed["receiving_record_id"].dropna().astype(str))
    if not ids:
        return
    times = changed.drop_duplicates("receiving_record_id").set_index("receiving_record_id")["receiving_time"]
    try:
        active, closed, _ = load_pool(logger)
    except Exception:  # noqa: BLE001
        return
    touched = 0
    for frame in (active, closed):
        if frame.empty or "receiving_record_id" not in frame.columns:
            continue
        hit = frame["receiving_record_id"].astype("string").isin(ids)
        if not hit.any():
            continue
        frame.loc[hit, "receiving_time"] = frame.loc[hit, "receiving_record_id"].map(times)
        touched += int(hit.sum())
    if touched:
        save_pool(active, closed, logger, reason="receiving date correction")
        logger.info("已同步 Pool 中 %d 条收货时间；未改 kpi_counted 与历史快照。", touched)


def _mark_manifests_date(source_file: str, confirmed: bool, argument: str) -> None:
    from common import read_json, write_json

    for path in list(DIRS["working"].glob("*/manifest.json")) + list(DIRS["step1_output"].glob("*/manifest.json")):
        try:
            data = read_json(path)
        except Exception:  # noqa: BLE001
            continue
        files = data.get("receiving_date_pending_files") or data.get("archived_receiving_files") or []
        names = {str(x) for x in files}
        if source_file not in names and not any(str(x).endswith(source_file) for x in names):
            continue
        data["receiving_date_argument"] = argument
        data["receiving_date_confirmed"] = confirmed
        data["receiving_date_pending_files"] = [] if confirmed else list(files)
        write_json(path, data)
