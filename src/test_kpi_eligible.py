# -*- coding: utf-8 -*-
"""AliExpress eligibility: ORD by name/customer; MIA non-AGS. Does not touch the live Pool."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from kpi_rules import (  # noqa: E402
    aliexpress_no_ags_mia,
    aliexpress_ord_ineligible,
    apply_eligibility_exclusions,
    coerce_eligible,
    eligible_label,
    is_procarrier_customer,
    kpi_ineligible_mask,
    source_has_ags,
    split_removed_rows,
)


def _row(**kwargs):
    base = {
        "TNO": "T",
        "CUSTOMER": "AliExpress",
        "业务名称": "AliExpress",
        "platform_sheet": "AliExpress",
        "Hub": "MIA",
        "has_ags": "No",
        "Source": "Platform",
        "original_source": "Platform",
    }
    base.update(kwargs)
    return base


def run_checks() -> list[tuple[bool, str]]:
    out: list[tuple[bool, str]] = []

    def check(ok: bool, label: str) -> None:
        out.append((bool(ok), label))

    df = pd.DataFrame([
        _row(TNO="A1", has_ags="No", Hub="MIA", Source="Platform"),
        _row(TNO="A2", has_ags="No", Hub="ORD", Source="System"),
        _row(TNO="A3", has_ags="Yes", Hub="MIA", Source="AGS;Platform", original_source="AGS"),
        _row(TNO="A4", has_ags="No", Hub="LAX", Source="Platform"),
        _row(TNO="A5", CUSTOMER="SHEIN", 业务名称="SHEIN", platform_sheet=None,
             has_ags="No", Hub="MIA", Source="Platform"),
        _row(TNO="A6", has_ags="No", Hub="mia", Source="Platform"),
        _row(TNO="A7", has_ags=None, Hub="ORD", Source="No Ascan Pool",
             original_source="Platform"),
        _row(TNO="A8", has_ags=None, Hub="ORD", Source="AGS;No Ascan Pool",
             original_source="AGS"),
        _row(TNO="A9", has_ags="Yes", Hub="ORD", Source="AGS;Platform", original_source="AGS"),
        _row(TNO="A10", CUSTOMER="SHEIN", 业务名称="AliExpress", Hub="ORD",
             has_ags="Yes", Source="AGS"),
        _row(TNO="A11", CUSTOMER="AliExpress", 业务名称="SHEIN", Hub="ORD",
             has_ags="Yes", Source="AGS"),
        _row(TNO="A12", CUSTOMER="SHEIN", 业务名称="SHEIN", platform_sheet=None,
             Hub="ORD", has_ags="Yes", Source="AGS"),
        _row(TNO="P1", CUSTOMER="ProCarrier", 业务名称="ProCarrier", platform_sheet=None,
             Hub="DFW", has_ags="No", Source="System", original_source="System"),
        _row(TNO="P2", CUSTOMER="pro carrier", 业务名称="ProCarrier", platform_sheet=None,
             Hub="MIA", has_ags="No", Source="System", original_source="System"),
        _row(TNO="P3", CUSTOMER="Pro Carrier T11", 业务名称="Pro Carrier T11", platform_sheet=None,
             Hub="DFW", has_ags="Yes", Source="AGS", original_source="AGS"),
        _row(TNO="P4", CUSTOMER="ProCarrier T11", 业务名称="ProCarrier T11", platform_sheet=None,
             Hub="EWR", has_ags="Yes", Source="AGS", original_source="AGS"),
    ])
    hit = kpi_ineligible_mask(df)
    got = set(df.loc[hit, "TNO"])
    check(got == {"A1", "A2", "A6", "A7", "A8", "A9", "A10", "A11", "P1", "P2"},
          f"排除集合含 AliExpress 规则 + ProCarrier（实际 {sorted(got)}）")
    check(not bool(hit[df["TNO"] == "A3"].iloc[0]), "有 AGS 的 AliExpress MIA 仍可进 KPI")
    check(not bool(hit[df["TNO"] == "A4"].iloc[0]), "AliExpress LAX 仍可进 KPI")
    check(not bool(hit[df["TNO"] == "A5"].iloc[0]), "非 AliExpress 的 MIA 不排除")
    check(not bool(hit[df["TNO"] == "A12"].iloc[0]), "非 AliExpress 的 ORD 不排除")
    check(bool(hit[df["TNO"] == "A8"].iloc[0]), "AliExpress ORD 即使有 AGS 也不进 KPI")
    check(bool(aliexpress_ord_ineligible(df)[df["TNO"] == "A9"].iloc[0]),
          "业务名称且 CUSTOMER 都是 AliExpress、Hub=ORD 排除")
    check(bool(aliexpress_ord_ineligible(df)[df["TNO"] == "A10"].iloc[0]),
          "仅 业务名称=AliExpress 且 Hub=ORD 排除")
    check(bool(aliexpress_ord_ineligible(df)[df["TNO"] == "A11"].iloc[0]),
          "仅 CUSTOMER=AliExpress 且 Hub=ORD 排除")
    check(bool(aliexpress_no_ags_mia(df)[df["TNO"] == "A1"].iloc[0]),
          "AliExpress 非AGS MIA 仍排除")
    check(not bool(aliexpress_no_ags_mia(df)[df["TNO"] == "A9"].iloc[0]),
          "MIA 非AGS 规则不再覆盖 ORD")
    check(bool(source_has_ags(df)[df["TNO"] == "A8"].iloc[0]), "Source 含 AGS 的回退识别")
    check(not bool(source_has_ags(df)[df["TNO"] == "A7"].iloc[0]), "仅 Pool 标签不算 AGS")

    due = pd.Series(True, index=df.index)
    eligible = apply_eligibility_exclusions(due, df)
    check(bool(eligible[df["TNO"] == "A3"].iloc[0]), "排除规则只改 eligible，有 AGS 的 MIA 仍为 True")
    check(not bool(eligible[df["TNO"] == "A1"].iloc[0]), "排除规则把 A1 标为 False")
    check(not bool(eligible[df["TNO"] == "A9"].iloc[0]), "排除规则把有 AGS 的 ORD AliExpress 标为 False")
    check(len(eligible) == len(df), "不删行，只改 kpi_eligible")

    check(bool(is_procarrier_customer(df)[df["TNO"] == "P1"].iloc[0]),
          "CUSTOMER=ProCarrier 排除")
    check(bool(is_procarrier_customer(df)[df["TNO"] == "P2"].iloc[0]),
          "空格/大小写的 Pro Carrier 也排除")
    check(not bool(is_procarrier_customer(df)[df["TNO"] == "P3"].iloc[0]),
          "Pro Carrier T11 不排除")
    check(not bool(is_procarrier_customer(df)[df["TNO"] == "P4"].iloc[0]),
          "ProCarrier T11 不排除")
    check(not bool(eligible[df["TNO"] == "P1"].iloc[0]), "ProCarrier kpi_eligible=False")
    check(bool(eligible[df["TNO"] == "P3"].iloc[0]), "Pro Carrier T11 仍可进 KPI")

    pool_only = pd.DataFrame([{
        "TNO": "T008",
        "CUSTOMER": "Vinted US Domestic",
        "业务名称": "Vinted US Domestic",
        "Customer": pd.NA,
        "platform_sheet": None,
        "Hub": "MIA",
        "has_ags": "No",
        "Source": "System; No Ascan Pool",
        "original_source": "System",
    }])
    pool_hit = is_procarrier_customer(pool_only)
    pool_elig = apply_eligibility_exclusions(pd.Series(True, index=pool_only.index), pool_only)
    check(bool(pool_hit.notna().all()) and not bool(pool_hit.iloc[0]),
          "Pool 捞出缺 Customer 列时 ProCarrier 判断不是 NA")
    check(bool(pool_elig.notna().all()) and bool(pool_elig.iloc[0]),
          "Pool 捞出的 Vinted MIA 仍 kpi_eligible=True")

    from dc_bridge import merge_module
    sla_df = pd.DataFrame({
        "CUSTOMER": ["ProCarrier", "Pro Carrier T11", "Vinted US Domestic"],
        "业务名称": ["ProCarrier", "Pro Carrier T11", "Vinted US Domestic"],
        "Source": ["System", "AGS", "System"],
        "has_system": ["Yes", "No", "Yes"],
        "has_ags": ["No", "Yes", "No"],
        "start_time": pd.to_datetime(["2026-09-01 10:00:00"] * 3),
        "Hub": ["DFW", "DFW", "DFW"],
    })
    sla48 = merge_module().is_48h_sla_row(sla_df)
    spec = merge_module().compute_sla_fields(sla_df, sla_df["start_time"])
    check(bool(sla48.iloc[0]) and int(spec["sla_hours"].iloc[0]) == 48,
          "System-only ProCarrier 仍是 48h SLA")
    check((not bool(sla48.iloc[1])) and int(spec["sla_hours"].iloc[1]) == 24,
          "Pro Carrier T11（AGS）仍是 24h，不改 T11")
    check(bool(sla48.iloc[2]) and int(spec["sla_hours"].iloc[2]) == 48,
          "System-only Vinted 仍是 48h")

    fbt_df = pd.DataFrame({
        "CUSTOMER": ["TT FBT", "TT FBT", "tt fbt", "TT GS"],
        "业务名称": ["TT FBT", "TT FBT", "FBT", "TT GS"],
        "platform_sheet": ["FBT", "FBT", "FBT", "TT-GS"],
        "Hub": ["MIA", "LAX", " mia ", "MIA"],
        "sla_hours": [24, 12, 12, 24],
        "start_time": pd.to_datetime(["2026-10-01 10:00:00"] * 4),
    })
    fbt_spec = merge_module().compute_sla_fields(fbt_df, fbt_df["start_time"])
    check(int(fbt_spec["sla_hours"].iloc[0]) == 12, "MIA TT FBT 是 12h")
    check(int(fbt_spec["sla_hours"].iloc[1]) == 24,
          "非 MIA TT FBT 是 24h，不沿用已存的 12h")
    check(int(fbt_spec["sla_hours"].iloc[2]) == 12, "Hub 大小写和空格后的 MIA TT FBT 仍是 12h")
    check(int(fbt_spec["sla_hours"].iloc[3]) == 24, "MIA 的非 FBT 仍是 24h")
    mia_due = pd.Timestamp("2026-10-01 22:00:00")
    lax_due = pd.Timestamp("2026-10-02 10:00:00")
    check(pd.Timestamp(fbt_spec["due_display"].iloc[0]) == mia_due, "MIA TT FBT 到期为起算 +12h")
    check(pd.Timestamp(fbt_spec["due_display"].iloc[1]) == lax_due, "其他 Hub TT FBT 到期为起算 +24h")

    drop_df = pd.DataFrame({
        "TNO": ["PC", "T11", "VIN", "IAH_HO", "IAH_PU", "DFW_FBT", "IAH_FBT_OLD", "SHEIN"],
        "CUSTOMER": ["ProCarrier", "Pro Carrier T11", "Vinted US Domestic", "SHEIN", "SHEIN",
                     "TT FBT", "TT FBT", "SHEIN"],
        "业务名称": ["ProCarrier", "Pro Carrier T11", "Vinted US Domestic", "SHEIN", "SHEIN",
                     "TT FBT", "TT FBT", "SHEIN"],
        "platform_sheet": [None, None, None, None, None, "FBT", "FBT", None],
        "Hub": ["DFW", "DFW", "MIA", "IAH", "IAH", "DFW", "IAH", "LAX"],
        "start_time_type": ["Handover Time", "Handover Time", "Handover Time", "Handover Time",
                            "Pickup Time", "Manifest Time", "Manifest Time", "Handover Time"],
        "merge_note": [None, None, None, None, None, None, "fbt_dfw_to_iah", None],
    })
    kept, removed = split_removed_rows(drop_df)
    reason = dict(zip(removed["TNO"], removed["删除原因"]))
    check(set(kept["TNO"]) == {"T11", "IAH_PU", "SHEIN", "IAH_FBT_OLD"},
          f"留下 T11、IAH Pickup、SHEIN、IAH TT FBT（实际 {sorted(kept['TNO'])}）")
    check(reason.get("PC") == "ProCarrier", "ProCarrier 删除")
    check(reason.get("VIN") == "Vinted", "Vinted 删除")
    check(reason.get("IAH_HO") == "IAH 且 Handover", "IAH Handover 删除")
    check(reason.get("DFW_FBT") == "DFW 的 TT FBT", "当前 Hub 仍是 DFW 的 TT FBT 删除")
    check("IAH_FBT_OLD" not in reason, "已改成 IAH 的 TT FBT 保留")
    check("T11" not in reason, "Pro Carrier T11 不删")

    labels = eligible_label(eligible)
    check(set(labels.unique()) <= {"True", "False"}, "报表标签只有 True/False")
    check(coerce_eligible(pd.Series(["True", "False", "Yes", "No"])).tolist()
          == [True, False, True, False], "coerce 识别 True/Yes")
    return out


if __name__ == "__main__":
    results = run_checks()
    failed = [label for ok, label in results if not ok]
    print(f"{sum(ok for ok, _ in results)}/{len(results)} passed")
    for ok, label in results:
        print(f"  [{'PASS' if ok else 'FAIL'}] {label}")
    raise SystemExit(1 if failed else 0)
