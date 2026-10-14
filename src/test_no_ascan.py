# -*- coding: utf-8 -*-
"""No-A-scan subclasses: empty code, 51005, other codes."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import yaml

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from kpi_rules import subclass_no_ascan  # noqa: E402


def run_checks() -> list[tuple[bool, str]]:
    out: list[tuple[bool, str]] = []

    def check(ok: bool, label: str) -> None:
        out.append((bool(ok), label))

    cfg_path = HERE.parent / "config" / "status_check.yaml"
    cfg = yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
    status = pd.Series(["__NO_ASCAN__"] * 6)
    queried = pd.Series([False, True, True, True, True, True])
    event = pd.Series([None, "", "51005", "51005.0", "12345", "57407"])
    got = subclass_no_ascan(status, queried, event, cfg).tolist()
    check(got[0] == "No A-Scan (未查询)", f"未查询仍是未查询（实际 {got[0]}）")
    check(got[1] == "No Ascan No Track", f"事件码为空 -> No Ascan No Track（实际 {got[1]}）")
    check(got[2] == "51005", f"51005 -> 状态 51005（实际 {got[2]}）")
    check(got[3] == "51005", f"51005.0 -> 状态 51005（实际 {got[3]}）")
    check(got[4] == "Missing Ascan", f"其他事件码 -> Missing Ascan（实际 {got[4]}）")
    check(got[5] == "Missing Ascan", f"57407 无 A-scan 仍先标 Missing Ascan（实际 {got[5]}）")

    from kpi_rules import apply_wrong_address, wrong_address_match

    idx = pd.RangeIndex(5)
    wa_fields = {
        "last_event_code": pd.Series([None, None, None, None, None], index=idx),
        "last_event_desc": pd.Series([None, None, None, None, None], index=idx),
        "last_status_event_desc": pd.Series([None, None, None, None, None], index=idx),
        "last_attempt_event_desc": pd.Series([None, None, None, None, None], index=idx),
        "last_public_status_event_desc": pd.Series([None, None, None, None, None], index=idx),
        "arrived_sc_date_57407": pd.Series(
            ["2026-09-20 08:00:00", None, "", "  ", "nan"], index=idx
        ),
    }
    hit, src = wrong_address_match(wa_fields, cfg)
    check(bool(hit.iloc[0]), "Arrived Shipping Center Date57407 非空 -> Wrong Address")
    check("Arrived Shipping Center Date57407" in str(src.iloc[0]),
          f"命中来源写该列名（实际 {src.iloc[0]!r}）")
    check(not bool(hit.iloc[1]), "该列空白不命中")
    check(not bool(hit.iloc[2]), "该列空字符串不命中")
    check(not bool(hit.iloc[3]), "该列纯空格不命中")
    check(not bool(hit.iloc[4]), "该列 nan 文本不命中")
    status = pd.Series(["Late(Handover)", "On Time", "No Ascan No Track",
                        "Late(PickUp)", "Missing Ascan"], index=idx)
    new_status, blocked = apply_wrong_address(status, hit, cfg)
    check(new_status.iloc[0] == "Wrong Address", "Late 可被 Date57407 覆盖为 Wrong Address")
    on_time_hit = pd.Series([False, True, False, False, False], index=idx)
    kept, blocked_on_time = apply_wrong_address(status, on_time_hit, cfg)
    check(kept.iloc[1] == "On Time" and bool(blocked_on_time.iloc[1]),
          "On Time 即使 Date57407 非空也不自动覆盖")

    from kpi_rules import split_late

    late = pd.Series(["Late", "Late", "Late", "On Time", "Late(PickUp)", "Late(Others)", "Late"])
    types = pd.Series([
        "Pickup Time", "Handover Time", "Manifest Time", "Pickup Time",
        "Handover Time", "date_manifested", None,
    ])
    split = split_late(late, types, cfg).tolist()
    check(split[0] == "Late(PickUp)", f"Pickup Time -> Late(PickUp)（实际 {split[0]}）")
    check(split[1] == "Late(Handover)", f"Handover Time -> Late(Handover)（实际 {split[1]}）")
    check(split[2] == "Late(Manifest)", f"Manifest Time -> Late(Manifest)（实际 {split[2]}）")
    check(split[3] == "On Time", f"On Time 不拆 Late（实际 {split[3]}）")
    check(split[4] == "Late(PickUp)", f"已拆的 Late(PickUp) 保持（实际 {split[4]}）")
    check(split[5] == "Late(Manifest)", f"旧 Late(Others)+date_manifested -> Late(Manifest)（实际 {split[5]}）")
    check(split[6] == "Late(Others)", f"空 start_time_type 仍 Late(Others)（实际 {split[6]}）")
    return out


if __name__ == "__main__":
    failed = 0
    for ok, label in run_checks():
        print(f"  [{'PASS' if ok else 'FAIL'}] {label}")
        failed += int(not ok)
    raise SystemExit(failed)
