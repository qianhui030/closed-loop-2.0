# -*- coding: utf-8 -*-
"""System Hub: first_ascan_hub, else origin_poe; drop when both empty."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from dc_bridge import merge_module  # noqa: E402


def run_checks() -> list[tuple[bool, str]]:
    out: list[tuple[bool, str]] = []

    def check(ok: bool, label: str) -> None:
        out.append((bool(ok), label))

    m = merge_module()
    df = pd.DataFrame({
        "first_ascan_hub": ["LAX", None, "  ", None, "ORD"],
        "origin_poe": ["EWR", "DFW", "MIA", None, "IAH"],
    })
    got = m.pick_system_hub(df)
    check(str(got.iloc[0]) == "LAX", f"有 first_ascan_hub 用它（实际 {got.iloc[0]}）")
    check(str(got.iloc[1]) == "DFW", f"first_ascan_hub 空则 origin_poe（实际 {got.iloc[1]}）")
    check(str(got.iloc[2]) == "MIA", f"空白 first_ascan_hub 回退 origin_poe（实际 {got.iloc[2]}）")
    check(pd.isna(got.iloc[3]), f"两列都空则为空（实际 {got.iloc[3]}）")
    check(str(got.iloc[4]) == "ORD", f"两列都有时仍用 first_ascan_hub（实际 {got.iloc[4]}）")

    old = pd.DataFrame({"origin_poe": ["EWR", None]})
    old_got = m.pick_system_hub(old)
    check(str(old_got.iloc[0]) == "EWR", f"没有 first_ascan_hub 列时用 origin_poe（实际 {old_got.iloc[0]}）")
    idx = pd.Index(["t1", "t2", "t3", "t4"])
    ascan_s = pd.Series(["LAX", pd.NA, pd.NA, pd.NA], index=idx, dtype="object")
    plt_s = pd.Series(["EWR", "DFW", pd.NA, pd.NA], index=idx, dtype="object")
    sys_s = pd.Series(["LAX", "MIA", "ORD", pd.NA], index=idx, dtype="object")
    ags_s = pd.Series(["IAH", "IAH", "IAH", "ATL"], index=idx, dtype="object")
    hub, src = m.pick_master_hub(ascan_s, plt_s, sys_s, ags_s)
    check(str(hub.iloc[0]) == "LAX" and str(src.iloc[0]) == "System",
          f"first_ascan_hub 盖过 Platform（实际 hub={hub.iloc[0]} src={src.iloc[0]}）")
    check(str(hub.iloc[1]) == "DFW" and str(src.iloc[1]) == "Platform",
          f"无 first_ascan_hub 时 Platform 仍优先 origin_poe（实际 hub={hub.iloc[1]} src={src.iloc[1]}）")
    check(str(hub.iloc[2]) == "ORD" and str(src.iloc[2]) == "System",
          f"Platform 空则用 System origin_poe（实际 hub={hub.iloc[2]} src={src.iloc[2]}）")
    check(str(hub.iloc[3]) == "ATL" and str(src.iloc[3]) == "AGS",
          f"前三源都空才用 AGS（实际 hub={hub.iloc[3]} src={src.iloc[3]}）")

    fbt = pd.DataFrame({
        "CUSTOMER": ["TT FBT", "tt fbt", "TT FBT", "TT GS", "Winit", "TT FBT"],
        "业务名称": ["FBT", "something", "FBT", "TT GS", "Winit", "FBT"],
        "platform_sheet": ["FBT", "FBT", "FBT", "TT-GS", "Winit", "FBT"],
        "Hub": ["DFW", "dfw ", "IAH", "DFW", "DFW", "ORD"],
        "merge_note": ["sys", "", "", "", "", ""],
        "跨Hub转运": ["LAX-DFW", "", "", "", "", ""],
    })
    remapped, n = m.remap_fbt_dfw_hub(fbt)
    check(n == 2, f"TT FBT DFW 改 IAH 共 2 条（实际 {n}）")
    check(str(remapped.iloc[0]["Hub"]) == "IAH",
          f"CUSTOMER=TT FBT 且 Hub=DFW 改为 IAH（实际 {remapped.iloc[0]['Hub']}）")
    check(str(remapped.iloc[1]["Hub"]) == "IAH",
          f"大小写/空格的 DFW TT FBT 也改 IAH（实际 {remapped.iloc[1]['Hub']}）")
    check(str(remapped.iloc[2]["Hub"]) == "IAH",
          f"本来就是 IAH 的 TT FBT 不动（实际 {remapped.iloc[2]['Hub']}）")
    check(str(remapped.iloc[3]["Hub"]) == "DFW",
          f"TT GS 在 DFW 保持 DFW（实际 {remapped.iloc[3]['Hub']}）")
    check(str(remapped.iloc[4]["Hub"]) == "DFW",
          f"非 FBT 的 DFW 不动（实际 {remapped.iloc[4]['Hub']}）")
    check(str(remapped.iloc[5]["Hub"]) == "ORD",
          f"TT FBT 其他 Hub 不动（实际 {remapped.iloc[5]['Hub']}）")
    check(str(remapped.iloc[0]["跨Hub转运"]) == "LAX-DFW",
          "只改 Hub，不改 跨Hub转运 标签")
    check("fbt_dfw_to_iah" in str(remapped.iloc[0]["merge_note"]),
          f"merge_note 记录 fbt_dfw_to_iah（实际 {remapped.iloc[0]['merge_note']}）")
    again, n2 = m.remap_fbt_dfw_hub(remapped)
    check(n2 == 0, f"已经改过的不再重复改（实际 {n2}）")
    return out


if __name__ == "__main__":
    results = run_checks()
    failed = [label for ok, label in results if not ok]
    for ok, label in results:
        print(("OK  " if ok else "FAIL"), label)
    if failed:
        raise SystemExit(f"{len(failed)} failed")
    print(f"{len(results)} checks passed")
