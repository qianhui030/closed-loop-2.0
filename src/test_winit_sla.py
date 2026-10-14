# -*- coding: utf-8 -*-
"""Winit calendar D0/D1 SLA boundary cases. Does not touch the live Pool."""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from kpi_rules import classify_base_status, compute_kpi_times, on_time_count, time_diff_hours  # noqa: E402


def _merge_module():
    path = ROOT / "Data Combination" / "build_unified_master.py"
    spec = importlib.util.spec_from_file_location("build_unified_master_winit_test", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _row(tno, hub, start, ascan, customer="Winit", sheet="Winit"):
    return {
        "TNO": tno,
        "CUSTOMER": customer,
        "Customer": customer,
        "业务名称": customer,
        "platform_sheet": sheet,
        "Hub": hub,
        "start_time": start,
        "first_ascan_time": ascan,
        "status": None,
    }


def _frame(rows):
    return pd.DataFrame(rows)


def run_checks(verbose: bool = False) -> list[tuple[bool, str]]:
    results: list[tuple[bool, str]] = []

    def check(ok: bool, label: str) -> None:
        results.append((bool(ok), label))
        if verbose:
            print(f"  [{'PASS' if ok else 'FAIL'}] {label}", flush=True)

    bum = _merge_module()
    cases = [
        ("D1 over 24h", "JFK", "2026-08-22 08:00:00", "2026-08-23 23:00:00",
         "2026-08-23", "On Time", "calendar_d1"),
        ("D1 last second", "ORD", "2026-08-22 08:00:00", "2026-08-23 23:59:59",
         "2026-08-23", "On Time", "calendar_d1"),
        ("D1 past midnight", "ORD", "2026-08-22 08:00:00", "2026-08-24 00:00:00",
         "2026-08-23", "Late", "calendar_d1"),
        ("D0 same day", "LAX自提", "2026-08-23 08:00:00", "2026-08-23 23:00:00",
         "2026-08-23", "On Time", "calendar_d0"),
        ("D0 next day", "LAX自提", "2026-08-23 23:00:00", "2026-08-24 00:01:00",
         "2026-08-23", "Late", "calendar_d0"),
        ("KY Friday-Monday no skip", "LAX", "2026-08-21 10:00:00", "2026-08-24 10:00:00",
         "2026-08-22", "Late", "calendar_d1"),
        ("KY Sunday-Tuesday no skip", "LAX", "2026-08-23 10:00:00", "2026-08-25 10:00:00",
         "2026-08-24", "Late", "calendar_d1"),
    ]
    rows = [_row(f"W{i:03d}", hub, start, ascan)
            for i, (_, hub, start, ascan, *_) in enumerate(cases, 1)]
    rows.append(_row("WNA1", "JFK", "2026-08-22 08:00:00", None))
    rows.append(_row("WNA2", "JFK", None, "2026-08-23 10:00:00"))
    rows.append(_row("WLAX", "LAX", "2026-08-23 08:00:00", "2026-08-23 23:00:00"))
    df = _frame(rows)
    out = bum.add_ascan_due(df)
    times = compute_kpi_times(df, df["start_time"])

    for i, (label, hub, start, ascan, due, expected, sla_type) in enumerate(cases, 1):
        tno = f"W{i:03d}"
        row = out[out["TNO"] == tno].iloc[0]
        tr = times.loc[out["TNO"] == tno].iloc[0]
        check(str(row["status"]) == expected, f"{label}: status={row['status']} expected {expected}")
        check(str(tr["kpi_date"]) == due, f"{label}: kpi_date={tr['kpi_date']} expected {due}")
        check(str(row["sla_type"]) == sla_type, f"{label}: sla_type={row['sla_type']}")
        due_disp = pd.Timestamp(row["Ascan Due Date"])
        check(due_disp.strftime("%H:%M:%S") == "23:59:59",
              f"{label}: display due is 23:59:59 ({due_disp})")
        if ascan and expected == "On Time":
            diff = float(time_diff_hours(pd.Series([ascan]), pd.Series([start])).iloc[0])
            check(pd.isna(row["sla_hours"]), f"{label}: sla_hours is NA (not hour SLA)")
            if "over 24h" in label or "last second" in label:
                check(diff > 24, f"{label}: time_diff {diff} remains actual hours")

    na = out[out["TNO"] == "WNA1"].iloc[0]
    check(str(na["status"]) == "No A-Scan", f"no A-scan → No A-Scan (actual {na['status']})")
    missing = out[out["TNO"] == "WNA2"].iloc[0]
    check(pd.isna(missing["status"]) or str(missing["status"]) not in {"On Time", "Late"},
          f"missing Manifested Time is not defaulted to On Time/Late (actual {missing['status']})")
    check(pd.isna(times.loc[out["TNO"] == "WNA2", "kpi_date"].iloc[0]),
          "missing Manifested Time → no kpi_date")

    lax = out[out["TNO"] == "WLAX"].iloc[0]
    check(str(lax["sla_type"]) == "calendar_d1",
          f"ordinary Winit LAX stays D1, not 自提 D0 (actual {lax['sla_type']})")
    check(str(times.loc[out["TNO"] == "WLAX", "kpi_date"].iloc[0]) == "2026-08-24",
          "ordinary Winit LAX 8/23 manifest → 8/24 KPI")

    base = classify_base_status(
        out["first_ascan_time"], time_diff_hours(out["first_ascan_time"], out["start_time"]),
        out["sla_hours"], sla_type=out["sla_type"], kpi_due_date=out["Ascan Due Date"],
        kpi_deadline=times["kpi_deadline"],
    )
    ot = on_time_count(base.mask(base.astype("string") == "__NO_ASCAN__", "No A-Scan"))
    check(int(ot.sum()) == int((out["status"] == "On Time").sum()),
          "onTime_count matches add_ascan_due On Time count")
    check(bool((ot == (out["status"] == "On Time").astype("int64")).all()),
          "onTime_count row-aligned with status")

    hour = _frame([_row("H001", "DFW", "2026-08-22 08:00:00", "2026-08-23 07:00:00",
                        customer="SHEIN", sheet=None)])
    hour.loc[0, "业务名称"] = "SHEIN"
    hour.loc[0, "CUSTOMER"] = "SHEIN"
    hour.loc[0, "platform_sheet"] = None
    hour_out = bum.add_ascan_due(hour)
    check(str(hour_out.iloc[0]["sla_type"]) == "hours", "non-Winit remains hour SLA")
    check(int(hour_out.iloc[0]["sla_hours"]) == 24, "non-Winit default 24h")
    check(str(hour_out.iloc[0]["status"]) == "On Time", "hour SLA still uses start+24h")
    return results


def main() -> int:
    print("== Winit D0/D1 SLA 边界 ==")
    results = run_checks(verbose=True)
    passed = sum(1 for ok, _ in results if ok)
    print(f"Winit SLA 单测：{passed}/{len(results)} 通过")
    for ok, label in results:
        if not ok:
            print(f"  FAIL {label}")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
