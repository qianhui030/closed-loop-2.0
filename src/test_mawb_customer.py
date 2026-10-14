# -*- coding: utf-8 -*-
"""MAWB summary customer comes from CUSTOMER, not customer_raw."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from mawb_log import summarize_mawbs  # noqa: E402


def run_checks() -> list[tuple[bool, str]]:
    out: list[tuple[bool, str]] = []

    def check(ok: bool, label: str) -> None:
        out.append((bool(ok), label))

    df = pd.DataFrame({
        "TNO": ["A", "B"],
        "MAWB": ["003-31904493", "003-31904493"],
        "CUSTOMER": ["AliExpress", pd.NA],
        "customer_raw": ["Albatross America INC;AliExpress", "Albatross America INC"],
        "业务名称": ["AliExpress", "AGS"],
        "Hub": ["ORD", "ORD"],
        "Source": ["AGS;Platform", "AGS"],
        "start_time": ["2026-09-22 00:10:00", "2026-09-22 10:49:07"],
        "status": ["On Time", "No A-Scan"],
        "onTime_count": [1, 0],
    })
    got = summarize_mawbs(df, "2026-09-22", ["ORD"])
    name = str(got.iloc[0]["CUSTOMER"])
    check(name == "AliExpress", f"有标准化 CUSTOMER 时不用 customer_raw（实际 {name}）")

    blank = df.copy()
    blank["CUSTOMER"] = pd.NA
    blank_name = str(summarize_mawbs(blank, "2026-09-22", ["ORD"]).iloc[0]["CUSTOMER"])
    check(blank_name == "(空客户)", f"CUSTOMER 为空时不回退 customer_raw（实际 {blank_name}）")

    mixed = pd.DataFrame({
        "TNO": ["A", "B", "C"],
        "MAWB": ["CMDUCHN3934830", "CMDUCHN3934830", "CMDUCHN3934830"],
        "CUSTOMER": ["Whaleco", "Whaleco", "Whaleco"],
        "Hub": ["ORD", "ORD", "ATL"],
        "Source": ["Platform", "Platform", "Platform"],
        "start_time": ["2026-09-29 01:00:00", "2026-09-29 01:00:00", "2026-09-29 02:00:00"],
        "status": ["No Ascan No Track", "On Time", "On Time"],
        "onTime_count": [0, 1, 1],
    })
    hubs = ["MIA", "DFW", "IAH", "EWR", "ATL", "ORD"]
    split = summarize_mawbs(mixed, "2026-09-30", hubs)
    pairs = set(zip(split["Hub"].astype(str), split["Count of TNO"].astype(int)))
    check(pairs == {("ATL", 1), ("ORD", 2)},
          f"同一 MAWB 在 ATL 和 ORD 各一行（实际 {pairs}）")
    return out
