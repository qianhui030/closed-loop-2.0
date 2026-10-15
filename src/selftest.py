# -*- coding: utf-8 -*-
"""End-to-end self test on a throwaway tree. Never touches the live Pool or outputs.

    python src/selftest.py
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
REAL_ROOT = HERE.parent
SANDBOX = REAL_ROOT / "Working" / "_selftest_root"

WIDE_COLS = [
    "Source", "File名称", "primary_source", "业务名称", "TNO", "MAWB", "CUSTOMER",
    "customer_raw", "Hub", "start_time", "start_time_type", "first_ascan_time",
    "Ascan Due Date", "cohort_date", "time_diff_hours", "sla_hours", "status",
    "跨Hub转运", "Customer", "customer_standardized", "交接时间", "交接时间类型",
    "Ascan时间", "within_24h", "monitor_flag", "merge_note", "source_files",
    "source_count", "has_ags", "has_system", "has_platform", "customer_source",
    "hub_source", "start_time_source", "ascan_source", "mawb_source",
    "mawb_mapping_conflict_flag", "platform_sheet", "ags_mawb_hub",
    "ags_tracking_origin_hub", "ags_destination_hub", "ags_destination_facility",
    "ags_handover_time", "ags_ascan_time", "ags_destination_hubs",
]

RESULTS: list[tuple[bool, str]] = []


def check(ok: bool, label: str) -> None:
    RESULTS.append((bool(ok), label))
    print(f"  [{'PASS' if ok else 'FAIL'}] {label}", flush=True)


def eligible_flag(value) -> bool:
    """kpi_eligible from parquet may be True / 'True' / <NA>."""
    return str(value).strip().lower() in {"true", "yes", "1"}


def row(tno, source, biz, customer, hub, mawb, start, ascan, sla, cross=None,
        has_ags="No", has_system="No", has_platform="No", sheet=None,
        start_type="Handover Time"):
    due = pd.Timestamp(start) + pd.Timedelta(hours=sla)
    status = "No A-Scan" if ascan is None else (
        "On Time" if pd.Timestamp(ascan) <= due else "Late")
    return {
        "Source": source, "File名称": "selftest", "primary_source": source,
        "业务名称": biz, "TNO": tno, "MAWB": mawb, "CUSTOMER": customer,
        "customer_raw": customer, "Hub": hub, "start_time": start,
        "start_time_type": start_type, "first_ascan_time": ascan,
        "Ascan Due Date": due, "cohort_date": str(pd.Timestamp(start).date()),
        "time_diff_hours": None, "sla_hours": sla, "status": status,
        "跨Hub转运": cross, "Customer": customer, "customer_standardized": customer,
        "交接时间": start, "交接时间类型": start_type, "Ascan时间": ascan,
        "within_24h": None, "monitor_flag": None, "merge_note": None,
        "source_files": "selftest", "source_count": 1, "has_ags": has_ags,
        "has_system": has_system, "has_platform": has_platform,
        "customer_source": source, "hub_source": source,
        "start_time_source": source, "ascan_source": source, "mawb_source": source,
        "mawb_mapping_conflict_flag": "No", "platform_sheet": sheet,
        "ags_mawb_hub": hub if has_ags == "Yes" else None,
        "ags_tracking_origin_hub": None, "ags_destination_hub": None,
        "ags_destination_facility": None, "ags_handover_time": None,
        "ags_ascan_time": None, "ags_destination_hubs": None,
    }


DAY1 = [
    # LAX-DFW cross-hub group sharing MAWB M1
    row("T001", "AGS", "AGS", "TIKTOK", "LAX", "111-11111111", "2026-09-01 10:00:00",
        "2026-09-01 12:00:00", 24, "LAX-DFW", has_ags="Yes"),
    row("T002", "AGS", "AGS", "TIKTOK", "LAX", "111-11111111", "2026-09-01 10:00:00",
        None, 24, "LAX-DFW", has_ags="Yes"),
    row("T003", "AGS", "AGS", "TIKTOK", "LAX", "111-11111111", "2026-09-01 10:00:00",
        None, 24, "LAX-DFW", has_ags="Yes"),
    row("T004", "AGS", "AGS", "TIKTOK", "LAX", "111-11111111", "2026-09-01 10:00:00",
        "2026-09-03 12:00:00", 24, "LAX-DFW", has_ags="Yes"),
    # plain System record, no ascan, will be counted on D1 then sit in the pool
    row("T005", "System", "SHEIN", "SHEIN", "DFW", None, "2026-09-01 10:00:00",
        None, 24, None, has_system="Yes"),
    row("T006", "Platform", "Winit", "Winit", "JFK", None, "2026-09-01 10:00:00",
        "2026-09-02 23:00:00", 24, None, has_platform="Yes", sheet="Winit"),
    # 48h SLA, due 09-03: not in D1 scope, enters the pool uncounted
    row("T007", "System", "Demo Shipper", "Demo Shipper", "DFW", "222-22222222",
        "2026-09-01 10:00:00", None, 48, None, has_system="Yes"),
    row("T008", "System", "Demo Shipper", "Demo Shipper", "MIA", None,
        "2026-09-01 10:00:00", None, 48, None, has_system="Yes"),
    # Same MAWB received twice on D2; must NOT pick first or last time.
    row("T009", "System", "Demo Shipper", "Demo Shipper", "DFW", "777-77777777",
        "2026-09-01 10:00:00", None, 48, None, has_system="Yes"),
    # ProCarrier row: physically removed in Step 1, audited in 删除说明 / 本次删除
    row("T099", "System", "ProCarrier", "ProCarrier", "DFW", None,
        "2026-09-01 10:00:00", None, 24, None, has_system="Yes"),
    # 48h SLA with A-scan already in: must stay in the pool until 09-03 KPI
    row("T010", "System", "Demo Shipper", "Demo Shipper", "DFW", None,
        "2026-09-01 10:00:00", "2026-09-01 12:00:00", 48, None, has_system="Yes"),
    # EWR-ORD: first day stays EWR-ORD; pulled on due date becomes EWR-ORD(Found)
    row("T011", "AGS", "AGS", "TIKTOK", "EWR", "444-44444444", "2026-09-01 10:00:00",
        None, 48, "EWR-ORD", has_ags="Yes"),
    row("T012", "AGS", "AGS", "TIKTOK", "LAX", "555-55555555", "2026-09-01 10:00:00",
        None, 48, "LAX-DFW", has_ags="Yes"),
    # AliExpress, no AGS, MIA/ORD: stay on the master with kpi_eligible=False
    row("T020", "Platform", "AliExpress", "AliExpress", "MIA", None, "2026-09-01 10:00:00",
        "2026-09-01 12:00:00", 24, None, has_platform="Yes", sheet="AliExpress"),
    row("T021", "Platform", "AliExpress", "AliExpress", "ORD", None, "2026-09-01 10:00:00",
        None, 24, None, has_platform="Yes", sheet="AliExpress"),
]
DAY2 = [
    row("T101", "AGS", "AGS", "TIKTOK", "MIA", "333-33333333", "2026-09-02 10:00:00",
        "2026-09-02 15:00:00", 24, None, has_ags="Yes"),
    row("T102", "AGS", "AGS", "TIKTOK", "MIA", "333-33333333", "2026-09-02 10:00:00",
        None, 24, None, has_ags="Yes"),
    row("T103", "System", "SHEIN", "SHEIN", "EWR", None, "2026-09-02 10:00:00",
        "2026-09-02 12:00:00", 24, None, has_system="Yes", start_type="Pickup Time"),
]
DAY3 = [
    row("T201", "AGS", "AGS", "TIKTOK", "ORD", None, "2026-09-03 10:00:00",
        "2026-09-03 15:00:00", 24, None, has_ags="Yes"),
]

RECEIVING_CSV = """Customs Broker Name,Client,First Mile Type(*),MAWB,Receiver(*),Actual Time of Arrival (ATA)(*),Pallet Quantity Unloaded,Carton Quantity Unloaded
PRO CARRIER SPX,,CUSTOM BROKER,222-22222222,TESTER,20;00,3,
WHALECO,,CUSTOM BROKER,"111-11111111
999-99999999",TESTER,21;10,5,
SF EXPRESS,,LOCAL PICK UP,777-77777777,TESTER,08;00,2,
SF EXPRESS,,LOCAL PICK UP,777-77777777,TESTER,09;30,2,
SHIPBOB,,LOCAL PICK UP,,TESTER,13;30,24,
WHALECO,,CUSTOM BROKER,HDUJGLX26EC02554,TESTER,13;50,10,
"""


def write_wide(rows: list[dict]) -> None:
    df = pd.DataFrame(rows)[WIDE_COLS]
    target = SANDBOX / "Data Combination" / "output" / "unified_master.csv"
    target.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(target, index=False, encoding="utf-8-sig")


def write_status_check(tnos, ascan_map, event_map, name="status.csv", extra=None) -> None:
    extra = extra or {}
    rows = []
    for tno in tnos:
        more = extra.get(tno, {})
        rows.append({
            "TNO": tno,
            "First Ascan Time": ascan_map.get(tno, ""),
            "Last Physical Event Code": event_map.get(tno, ("", ""))[0],
            "Last Physical Event Description": event_map.get(tno, ("", ""))[1],
            "Last Status Event Description": more.get("status_desc", ""),
            'Last "Attempt" Event Description': more.get("attempt_desc", ""),
            "Last Public Status Event Description": more.get("public_desc", ""),
        })
    target = SANDBOX / "Status_Check" / "Input" / name
    target.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_csv(target, index=False, encoding="utf-8-sig")


def run(script: str, *args) -> subprocess.CompletedProcess:
    env = dict(os.environ)
    env["CLOSEDLOOP_ROOT"] = str(SANDBOX)
    env["PYTHONPATH"] = str(HERE)
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.run(
        [sys.executable, str(HERE / script), *args],
        capture_output=True, text=True, encoding="utf-8", errors="replace", env=env,
    )
    tag = f"{script} {' '.join(args)}"
    if proc.returncode != 0:
        print(f"--- {tag} 失败 ---")
        print(proc.stdout[-4000:])
    return proc


def latest_manifest(kpi_date: str) -> dict:
    import json
    files = sorted((SANDBOX / "Working").glob("*_step1/manifest.json"))
    picked = [json.loads(p.read_text(encoding="utf-8")) for p in files]
    picked = [m for m in picked if m.get("kpi_date") == kpi_date]
    return picked[-1]


def read_pool() -> tuple[pd.DataFrame, pd.DataFrame]:
    active_path = SANDBOX / "No_Ascan_Pool" / "Active_Pool.parquet"
    closed_path = SANDBOX / "No_Ascan_Pool" / "Closed_History.parquet"
    if active_path.exists():
        active = pd.read_parquet(active_path)
        closed = pd.read_parquet(closed_path) if closed_path.exists() else pd.DataFrame()
        for frame in (active, closed):
            if "TNO" in frame.columns:
                frame["TNO"] = frame["TNO"].astype(str)
        return active, closed
    path = SANDBOX / "No_Ascan_Pool" / "No_Ascan_Pool.xlsx"
    book = pd.read_excel(path, sheet_name=None, dtype={"TNO": str})
    return book["Active_Pool"], book["Closed_History"]


def main() -> int:
    print("== Winit D0/D1 SLA 边界 ==")
    from test_winit_sla import run_checks
    for ok, label in run_checks():
        check(ok, label)
    print("== 不进分母：AliExpress ORD/MIA；ProCarrier（T11 仍计入）==")
    from test_kpi_eligible import run_checks as run_elig_checks
    for ok, label in run_elig_checks():
        check(ok, label)
    print("== 无 A-scan 三类：空码 / 51005 / 其他 ==")
    from test_no_ascan import run_checks as run_no_ascan_checks
    for ok, label in run_no_ascan_checks():
        check(ok, label)
    print("== System Hub / TT FBT DFW->IAH ==")
    from test_system_hub import run_checks as run_hub_checks
    for ok, label in run_hub_checks():
        check(ok, label)

    global SANDBOX
    if SANDBOX.exists():
        try:
            shutil.rmtree(SANDBOX)
        except OSError:
            SANDBOX = REAL_ROOT / "Working" / f"_selftest_root_{os.getpid()}"
    SANDBOX.mkdir(parents=True)
    shutil.copytree(REAL_ROOT / "config", SANDBOX / "config")
    (SANDBOX / "Data Combination").mkdir(parents=True, exist_ok=True)
    shutil.copy2(REAL_ROOT / "Data Combination" / "build_unified_master.py",
                 SANDBOX / "Data Combination" / "build_unified_master.py")
    (SANDBOX / "Receiving_Log" / "Input").mkdir(parents=True, exist_ok=True)
    print(f"沙箱：{SANDBOX}\n")

    # ---------------------------------------------------------------- day 1
    print("== D1  KPI 日期 2026-09-02，无收货文件 ==")
    write_wide(DAY1)
    proc = run("stage1.py", "--kpi-date", "2026-09-02", "--merge", "reuse")
    check(proc.returncode == 0, "D1 第一步运行成功")
    m1 = latest_manifest("2026-09-02")
    check(m1["counts"]["kpi_eligible"] == 6, f"D1 本次 KPI 范围 = 6（实际 {m1['counts']['kpi_eligible']}）")
    check(m1["counts"]["master_rows"] == 8, f"D1 主表含到期但不可计入的 AliExpress：8（实际 {m1['counts']['master_rows']}）")
    d1 = pd.read_parquet(SANDBOX / "Working" / m1["run_id"] / "unified_master_step1.parquet")
    t020 = d1[d1["TNO"] == "T020"]
    t021 = d1[d1["TNO"] == "T021"]
    check(len(t020) == 1 and not eligible_flag(t020.iloc[0]["kpi_eligible"]),
          "T020 AliExpress 非AGS MIA 留在主表且 kpi_eligible=False")
    check(len(t021) == 1 and not eligible_flag(t021.iloc[0]["kpi_eligible"]) and pd.isna(t021.iloc[0]["query_reason"]),
          "T021 AliExpress 非AGS ORD 不进查询清单")
    d1_cross = set(d1.loc[d1["TNO"].isin(["T001", "T002", "T003", "T004"]), "跨Hub转运"].astype(str))
    check(d1_cross == {"LAX-DFW"},
          f"D1 当天考核 LAX-DFW 不改标（实际 {d1_cross}）")
    check("T011" not in set(d1["TNO"].astype(str)) and "T012" not in set(d1["TNO"].astype(str)),
          "D1 未到期跨Hub 不进主表")
    d1_pool, _ = read_pool()
    check(str(d1_pool.loc[d1_pool["TNO"] == "T011", "跨Hub转运"].iloc[0]) == "EWR-ORD",
          "D1 未到期 EWR-ORD 进 Pool 且不改标")
    check(str(d1_pool.loc[d1_pool["TNO"] == "T012", "跨Hub转运"].iloc[0]) == "LAX-DFW",
          "D1 未到期 LAX-DFW 进 Pool 且不改标")

    write_status_check(
        ["T002", "T004", "T005", "T007"],
        {"T004": "2026-09-03 12:00:00"},
        {"T002": ("57407", "Undelivered - Invalid Address"),
         "T005": ("", ""),
         "T007": ("12345", "In transit")},
    )
    proc = run("stage2.py", "--run-id", m1["run_id"])
    check(proc.returncode == 0, "D1 第二步运行成功")

    report = SANDBOX / "Final_Output" / "2026-09-02" / m1["run_id"] / "Ascan_KPI.xlsx"
    check(report.exists(), "D1 生成 Ascan_KPI.xlsx")
    book = pd.read_excel(report, sheet_name=None, dtype={"TNO": str})
    check(all(s in book for s in ("Updated Unified Master", "删除说明", "本次删除",
                                  "Hub Pivot", "Customer Pivot", "On Time Rate")),
          "报表含 Updated Unified Master、删除说明、本次删除与三个 Pivot 工作表")
    check("MAWB Pivot" not in book and "MAWB 汇总" not in book, "报表不再包含 MAWB 表")
    rem_detail = book["本次删除"]
    check("T099" in set(rem_detail["TNO"].astype(str)),
          "ProCarrier T099 已物理删除并记录在 本次删除 页签")
    proc_reb = run("stage2.py", "--run-id", m1["run_id"], "--rebuild-report")
    check(proc_reb.returncode == 0, "D1 --rebuild-report 重出报表成功")
    book_reb = pd.read_excel(report, sheet_name=None, dtype={"TNO": str})
    check("T099" in set(book_reb["本次删除"]["TNO"].astype(str)),
          "--rebuild-report 重出报表保留历史已删除明细 T099")
    rem_sum_reb = dict(zip(book_reb["删除说明"]["项目"], book_reb["删除说明"]["条数"]))
    check(int(rem_sum_reb.get("ProCarrier", 0)) == 1 and int(rem_sum_reb.get("合计删除", 0)) == 1,
          f"--rebuild-report 删除说明统计准确（ProCarrier={rem_sum_reb.get('ProCarrier')}）")
    upd = book["Updated Unified Master"]
    check(len(upd) == 8 and upd["TNO"].nunique() == 8, "主表保留本次到期 8 个 TNO（含不可计入）")
    elig_upd = upd[upd["kpi_eligible"].astype(str).isin(["True", "TRUE"])]
    check(len(elig_upd) == 6, f"默认分母仍是 6（实际 {len(elig_upd)}）")
    check(str(upd.loc[upd["TNO"] == "T020", "kpi_eligible"].iloc[0]) in {"False", "FALSE"},
          "报表里 T020 kpi_eligible=False")
    check("Pivot_Source" not in book and "未来到期记录" not in book,
          "报表不再含 Pivot_Source / 未来到期记录")
    st = dict(zip(upd["TNO"], upd["kpi_status"]))
    check(st.get("T002") == "Wrong Address", f"T002 判为 Wrong Address（实际 {st.get('T002')}）")
    check(st.get("T003") == "No A-Scan (未查询)", f"T003 未查询保留（实际 {st.get('T003')}）")
    check(st.get("T005") == "No Ascan No Track", f"T005 事件码为空 -> No Ascan No Track（实际 {st.get('T005')}）")
    check(st.get("T004") == "Late(Handover)", f"T004 判为 Late(Handover)（实际 {st.get('T004')}）")
    check(st.get("T001") == "On Time", f"T001 判为 On Time（实际 {st.get('T001')}）")
    check(st.get("T006") == "On Time", f"T006 Winit D1 跨过24h仍 On Time（实际 {st.get('T006')}）")
    t001 = upd[upd["TNO"] == "T001"].iloc[0]
    check(float(t001["time_diff_hours"]) < float(t001["sla_hours"]),
          "T001 On Time 来自 time_diff_hours <= sla_hours")
    t006 = upd[upd["TNO"] == "T006"].iloc[0]
    check(float(t006["time_diff_hours"]) > 24,
          f"T006 实际耗时超过 24h（实际 {t006['time_diff_hours']}）")
    check(str(t006["sla_type"]) == "calendar_d1",
          f"T006 sla_type=calendar_d1（实际 {t006['sla_type']}）")
    check(int(t006["onTime_count"]) == 1, "T006 onTime_count 与 On Time 对齐")
    aligned = (upd["onTime_count"].astype("int64")
               == (upd["kpi_status"].astype("string") == "On Time").astype("int64"))
    check(bool(aligned.all()), "onTime_count 与最终 status 逐行一致")
    t002_src = str(upd.loc[upd["TNO"] == "T002", "wrong_address_source"].iloc[0])
    check("57407" in t002_src or "Last Physical Event Code" in t002_src,
          f"T002 Wrong Address 记录命中列（实际 {t002_src}）")

    active, closed = read_pool()
    pool_tnos = set(active["TNO"])
    check({"T002", "T003", "T005", "T007", "T008", "T009", "T010", "T011", "T012"} <= pool_tnos,
          f"未到期/未 A-scan 记录进入 Active_Pool（实际 {sorted(pool_tnos)}）")
    check("T021" not in pool_tnos, "不可计入的 AliExpress 不写回 Pool")
    counted = dict(zip(active["TNO"], active["kpi_counted"].astype(str)))
    check(counted.get("T007") != "Yes" and counted.get("T008") != "Yes"
          and counted.get("T010") != "Yes",
          "未来到期记录留在池中且未计入")
    t010_pool = active[active["TNO"] == "T010"].iloc[0]
    check(pd.notna(t010_pool["first_ascan_time"]),
          "未到期但已有 A-scan 的记录带 A-scan 进入 Pool")
    check("T010" not in set(upd["TNO"].astype(str)),
          "未到期已有 A-scan 不进主表，只在 Pool 等待")
    check(counted.get("T002") == "Yes", "本次已考核的未 A-scan 记录如实记录 kpi_counted=Yes")
    wa = dict(zip(active["TNO"], active["status"].astype(str)))
    check(wa.get("T002") == "Wrong Address", "Wrong Address 异常类型在池中保留，未退回普通 No A-scan")

    # ------------------------------------------------- duplicate receiving import
    print("\n== 收货重复导入 ==")
    inp = SANDBOX / "Receiving_Log" / "Input"
    (inp / "DFW Receiving log (902).csv").write_text(RECEIVING_CSV, encoding="utf-8-sig")
    write_wide(DAY2)
    proc = run("stage1.py", "--kpi-date", "2026-09-03", "--merge", "reuse",
               "--receiving-date", "2026-09-02")
    check(proc.returncode == 0, "D2 第一步运行成功")
    m2 = latest_manifest("2026-09-03")
    hist = pd.read_csv(SANDBOX / "Receiving_Log" / "Standardized" / "receiving_history.csv", dtype=str)
    first_n = len(hist)
    check(first_n == 7, f"收货历史 7 条（含无MAWB与非标准单号），实际 {first_n}")
    check((hist["match_status"] == "待核对_无MAWB").sum() == 1, "无 MAWB 行保留在历史中")
    check((hist["match_status"] == "待核对_非标准单号").sum() == 1, "非标准单号保留原值并标记待确认")
    check(hist.loc[hist["match_status"] == "待核对_非标准单号", "MAWB"].iloc[0] == "HDUJGLX26EC02554",
          "非标准单号未改写成航空主单格式")

    (inp / "DFW Receiving log (902) 副本.csv").write_text(RECEIVING_CSV, encoding="utf-8-sig")
    proc = run("stage1.py", "--kpi-date", "2026-09-03", "--merge", "reuse",
               "--receiving-date", "2026-09-02")
    check(proc.returncode == 0, "重复导入同一份收货记录运行成功")
    hist2 = pd.read_csv(SANDBOX / "Receiving_Log" / "Standardized" / "receiving_history.csv", dtype=str)
    check(len(hist2) == first_n, f"重复导入不产生重复事件：仍是 {first_n} 条（实际 {len(hist2)}）")

    step1 = pd.read_parquet(SANDBOX / "Working" / m2["run_id"] / "unified_master_step1.parquet")
    check(step1["TNO"].nunique() == len(step1), "三源与 Pool 同 TNO 只保留一行")
    check("T007" not in set(step1["TNO"].astype(str)),
          "收货后仍未到期的记录不进主表，留在 Pool")
    pool_d2, _ = read_pool()
    t007 = pool_d2[pool_d2["TNO"] == "T007"].iloc[0]
    check(str(t007["start_time"])[:19] == "2026-09-01 10:00:00",
          f"原 start_time 保留（实际 {t007['start_time']}）")
    check(str(t007["kpi_start_time"])[:19] == "2026-09-02 20:00:00",
          f"Pool 采用 receiving_time 作为 kpi_start_time（实际 {t007['kpi_start_time']}）")
    check(t007["kpi_start_time_source"] == "Receiving Log", "kpi_start_time_source = Receiving Log")
    check("No Ascan Pool" in str(t007["Source"]), f"Source 标记 No Ascan Pool（实际 {t007['Source']}）")
    t101 = step1[step1["TNO"] == "T101"].iloc[0]
    check(t101["进表来源"] == "当天三源",
          f"T101 当天三源进表（实际 {t101['进表来源']}）")
    t012 = step1[step1["TNO"] == "T012"].iloc[0]
    check(str(t012["跨Hub转运"]) == "LAX-DFW(Found)",
          f"到期捞出的 LAX-DFW 改标为 LAX-DFW(Found)（实际 {t012['跨Hub转运']}）")
    t011 = step1[step1["TNO"] == "T011"].iloc[0]
    check(str(t011["跨Hub转运"]) == "EWR-ORD(Found)",
          f"到期捞出的 EWR-ORD 改标为 EWR-ORD(Found)（实际 {t011['跨Hub转运']}）")
    check(str(t007["original_source"]) == "System", f"original_source 保留原值（实际 {t007['original_source']}）")
    gap = (pd.to_datetime(t007["receiving_time"]) - pd.to_datetime(t007["start_time"])) / pd.Timedelta(hours=1)
    check(abs(float(gap) - 34.0) < 0.01, f"receiving 与 start 相差 34 小时（实际 {gap}）")
    t008 = step1[step1["TNO"] == "T008"].iloc[0]
    check(eligible_flag(t008["kpi_eligible"]), "到期未 A-scan 的 Pool 记录照样进入本次 KPI")
    t010 = step1[step1["TNO"] == "T010"].iloc[0]
    check(eligible_flag(t010["kpi_eligible"]), "到期且已有 A-scan 的 Pool 记录照样进入本次 KPI")
    check(pd.notna(pd.to_datetime(t010["first_ascan_time"], errors="coerce")),
          "T010 从 Pool 带出原 A-scan")
    check("T005" not in set(step1["TNO"].astype(str)),
          "已计入且未出现在当天收货 Log 的 Pool 记录不整池并入主表")
    check("T005" in set(pool_d2["TNO"].astype(str)),
          "未捞出的 Pool 记录仍保留在 Active_Pool")
    t009 = step1[step1["TNO"] == "T009"].iloc[0]
    check(pd.isna(t009["receiving_time"]) or str(t009["receiving_time"]) in {"", "NaT", "None"},
          f"分批收货未擅自取第一笔或最后一笔（T009 receiving_time={t009['receiving_time']}）")
    conflict = pd.read_csv(SANDBOX / "Working" / m2["run_id"] / "already_counted_receiving_conflict.csv")
    check(set(conflict["TNO"]) == {"T002", "T003"},
          f"已计入历史 KPI 后确认收货 -> 单独输出（实际 {list(conflict['TNO'])}）")
    pend = pd.read_csv(SANDBOX / "Working" / m2["run_id"] / "receiving_pending_review.csv")
    t009_pend = pend[pend["TNO"] == "T009"]
    check(len(t009_pend) == 1 and "多个收货批次" in str(t009_pend.iloc[0]["reason"]),
          f"T009 因同一 MAWB 两笔收货进入待核对（实际 {t009_pend.to_dict('records')}）")
    print(f"     待核对收货 {len(pend)} 条")

    print("\n== 收货日期待确认会挡住正式第二步 ==")
    proc = run("stage2.py", "--run-id", m2["run_id"])
    check(proc.returncode != 0 and "待确认" in (proc.stdout + proc.stderr),
          "收货日期未确认时拒绝正式第二步")

    env = dict(os.environ)
    env["CLOSEDLOOP_ROOT"] = str(SANDBOX)
    env["PYTHONPATH"] = str(HERE)
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.run(
        [sys.executable, str(HERE / "receiving_tool.py"), "confirm",
         "--source-file", "DFW Receiving log (902).csv"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", env=env,
    )
    if proc.returncode != 0:
        print(proc.stdout[-2000:])
        print(proc.stderr[-2000:])
    check(proc.returncode == 0, "收货日期确认成功")

    (inp / "DFW Receiving log (902) 再导入.csv").write_text(RECEIVING_CSV, encoding="utf-8-sig")
    proc = run("stage1.py", "--kpi-date", "2026-09-03", "--merge", "reuse",
               "--receiving-date", "2026-09-03")
    check(proc.returncode == 0, "用不同日期再次导入同一收货文件运行成功")
    hist3 = pd.read_csv(SANDBOX / "Receiving_Log" / "Standardized" / "receiving_history.csv", dtype=str)
    check(len(hist3) == first_n, f"不同日期再导入不产生第二套事件：仍是 {first_n} 条（实际 {len(hist3)}）")
    times = pd.to_datetime(hist3["receiving_time"], errors="coerce")
    check(times.dt.strftime("%Y-%m-%d").nunique() == 1,
          "再导入未把历史改成两套日期；日期更正只能走 correct")

    write_status_check(["T101", "T102", "T008"], {"T101": "2026-09-02 15:00:00"},
                       {"T102": ("999", "In transit"), "T008": ("51005.0", "")},
                       extra={"T102": {"status_desc": "Address is unknown/Wrong"}})
    proc = run("stage2.py", "--run-id", m2["run_id"], "--inspect-only")
    check(proc.returncode == 0, "Status Check 覆盖率检查成功")
    inspect = json.loads((SANDBOX / "Working" / m2["run_id"] / "status_check_inspect.json").read_text(encoding="utf-8"))
    cov = inspect["coverage"]
    check(cov["should_query"] == inspect["coverage"]["should_query"]
          and cov["should_query"] == int((step1["query_reason"].notna()).sum()),
          f"覆盖率分母 = Status_Check_List {cov['should_query']}，不是完整主表 {len(step1)}")
    check(cov["master_tno_total"] == len(step1), "完整主表数量单独报告")
    print(f"     应查询 {cov['should_query']}  匹配 {cov['matched']}  缺失 {cov['missing']}  多余 {cov['extra']}")

    proc = run("stage2.py", "--run-id", m2["run_id"])
    check(proc.returncode == 0, "D2 第二步运行成功")
    d2_upd = pd.read_excel(
        SANDBOX / "Final_Output" / "2026-09-03" / m2["run_id"] / "Ascan_KPI.xlsx",
        sheet_name="Updated Unified Master", dtype={"TNO": str},
    )
    st2 = dict(zip(d2_upd["TNO"], d2_upd["kpi_status"]))
    check(st2.get("T102") == "Wrong Address",
          f"T102 由 Last Status Event Description 判为 Wrong Address（实际 {st2.get('T102')}）")
    check(st2.get("T008") == "51005",
          f"T008 事件码 51005 且无 A-scan -> 51005（实际 {st2.get('T008')}）")
    check(st2.get("T010") == "On Time",
          f"T010 未到期已有 A-scan，到期日计入 On Time（实际 {st2.get('T010')}）")

    print("\n== 重跑幂等 ==")
    proc = run("stage2.py", "--run-id", m2["run_id"])
    check(proc.returncode != 0 and "不是待处理" in (proc.stdout + proc.stderr),
          "已完成批次再跑第二步会被拒绝，不重复计数")
    active2, closed2 = read_pool()
    t007_rows = active2[active2["TNO"] == "T007"]
    check(len(t007_rows) <= 1, "Active_Pool 中同一 TNO 不重复追加")
    check("T010" not in set(active2["TNO"].astype(str)),
          "已有 A-scan 且已计入的记录离开 Active_Pool")
    check("T010" in set(closed2["TNO"].astype(str)),
          "已有 A-scan 且已计入的记录进入 Closed_History")

    # ---------------------------------------------------------------- day 3
    print("\n== D3  KPI 日期 2026-09-04，无任何新收货文件 ==")
    write_wide(DAY3)
    proc = run("stage1.py", "--kpi-date", "2026-09-04", "--merge", "reuse")
    check(proc.returncode == 0, "D3 第一步运行成功")
    m3 = latest_manifest("2026-09-04")
    s3 = pd.read_parquet(SANDBOX / "Working" / m3["run_id"] / "unified_master_step1.parquet")
    t007 = s3[s3["TNO"] == "T007"]
    check(len(t007) == 1 and eligible_flag(t007.iloc[0]["kpi_eligible"]),
          "未来到期的 Pool 记录，在没有新收货文件的当天仍自动转入")

    # -------------------------------------------------------------- pivot check
    # Pivot check is only meaningful if D2 actually published a report.
    print("\n== 三个 Pivot 口径一致 ==")
    d2_report = SANDBOX / "Final_Output" / "2026-09-03" / m2["run_id"] / "Ascan_KPI.xlsx"
    d2_mia = SANDBOX / "Final_Output" / "2026-09-03" / m2["run_id"] / "Ascan_KPI_MIA.xlsx"
    if not d2_report.exists():
        check(False, "D2 未生成报表，跳过 Pivot 核对")
        print("\n" + "=" * 60)
        passed = sum(1 for ok, _ in RESULTS if ok)
        print(f"自测结果：{passed}/{len(RESULTS)} 通过")
        for ok, label in RESULTS:
            if not ok:
                print(f"  FAIL {label}")
        print("=" * 60)
        return 0 if passed == len(RESULTS) else 1
    src = pd.read_excel(d2_report, sheet_name="Updated Unified Master", dtype={"TNO": str})
    elig_src = src[src["kpi_eligible"].astype(str).isin(["True", "TRUE"])]
    overall = elig_src["onTime_count"].mean()
    hub_rate = elig_src.groupby("Hub")["onTime_count"].mean().mean()
    check(elig_src["TNO"].nunique() == len(elig_src), "默认分母是本次 kpi_eligible=True 的唯一 TNO")
    check(abs(overall - elig_src["onTime_count"].sum() / len(elig_src)) < 1e-9,
          f"总率按包裹数加权 = {overall:.4f}")
    print(f"     各 Hub 百分比简单平均 = {hub_rate:.4f}（与加权总率不同属正常，报表用加权）")
    check(not d2_mia.exists(), "D2 默认不再生成 Ascan_KPI_MIA.xlsx")
    try:
        import win32com.client as client

        def page_field(pt, field_name: str):
            for field in pt.PageFields:
                name = str(field.Name)
                try:
                    source = str(field.SourceName)
                except Exception:
                    source = name
                if name == field_name or source == field_name:
                    return field
            return pt.PivotFields(field_name)

        def page_value(pt, field_name: str) -> str:
            return str(page_field(pt, field_name).CurrentPage)

        def page_visible(pt, field_name: str) -> set[str]:
            field = page_field(pt, field_name)
            try:
                multi = bool(field.EnableMultiplePageItems)
            except Exception:
                multi = False
            if not multi:
                return {str(field.CurrentPage)}
            visible = set()
            for item in field.PivotItems():
                try:
                    if item.Visible:
                        visible.add(str(item.Name))
                except Exception:
                    continue
            return visible

        app = client.DispatchEx("Excel.Application")
        try:
            app.Visible = False
            app.DisplayAlerts = False
        except Exception:
            pass
        wb = app.Workbooks.Open(str(d2_report), UpdateLinks=0, ReadOnly=True)
        totals = {}
        hub_ws = wb.Worksheets("Hub Pivot")
        hub_names = [pt.Name for pt in hub_ws.PivotTables()]
        check(set(hub_names) == {"HubHandoverPct", "HubHandoverQty", "HubPickupPct", "HubPickupQty"},
              f"Hub Pivot 含 4 张表（实际 {sorted(hub_names)}）")
        split_names = (
            "HubHandoverPct", "HubHandoverQty", "HubPickupPct", "HubPickupQty",
            "OnTimeHandoverPct", "OnTimeHandoverQty", "OnTimePickupPct", "OnTimePickupQty",
        )
        ontime_ws = wb.Worksheets("On Time Rate")
        ontime_names = [pt.Name for pt in ontime_ws.PivotTables()]
        check(set(ontime_names) == {"OnTimeHandoverPct", "OnTimeHandoverQty",
                                    "OnTimePickupPct", "OnTimePickupQty"},
              f"On Time Rate 含 4 张表（实际 {sorted(ontime_names)}）")
        for name in split_names:
            ws = hub_ws if name.startswith("Hub") else ontime_ws
            pt = ws.PivotTables(name)
            totals[f"{name}_cross"] = page_value(pt, "跨Hub转运")
            totals[f"{name}_entry"] = page_value(pt, "进表来源")
            totals[f"{name}_elig"] = page_value(pt, "kpi_eligible")
            totals[f"{name}_start"] = page_value(pt, "start_time_type")
            totals[f"{name}_start_items"] = page_visible(pt, "start_time_type")
        pt = wb.Worksheets("Customer Pivot").PivotTables("CustomerPivot")
        totals["Customer Pivot"] = page_value(pt, "跨Hub转运")
        totals["Customer Pivot_entry"] = page_value(pt, "进表来源")
        totals["Customer Pivot_elig"] = page_value(pt, "kpi_eligible")
        totals["Customer Pivot_rows"] = [str(f.Name) for f in pt.RowFields]
        for name in ("OnTimeHandoverPct", "OnTimeHandoverQty", "OnTimePickupPct", "OnTimePickupQty"):
            totals[f"{name}_rows"] = [str(f.Name) for f in ontime_ws.PivotTables(name).RowFields]
        if d2_mia.exists():
            mia_wb = app.Workbooks.Open(str(d2_mia), UpdateLinks=0, ReadOnly=True)
            visible = [ws.Name for ws in mia_wb.Worksheets if int(ws.Visible) != 0]
            hidden = [ws.Name for ws in mia_wb.Worksheets if int(ws.Visible) == 0]
            check(set(visible) >= {"Pivot", "Handover", "PickUp", "TTFBT"},
                  f"MIA 报表可见页含 Pivot / Handover / PickUp / TTFBT（实际 {visible}）")
            check("MIA Source" in hidden, f"MIA Source 隐藏（实际 hidden={hidden}）")
            mia_names = [pt.Name for pt in mia_wb.Worksheets("Pivot").PivotTables()]
            check(set(mia_names) == {"MIAHandoverRate", "MIAPickupRate"},
                  f"MIA Pivot 含 Handover / Pick Up 两张表（实际 {sorted(mia_names)}）")
            mia_ho = mia_wb.Worksheets("Pivot").PivotTables("MIAHandoverRate")
            mia_pu = mia_wb.Worksheets("Pivot").PivotTables("MIAPickupRate")
            totals["MIA_rows"] = [str(f.Name) for f in mia_ho.RowFields]
            totals["MIA_pickup_rows"] = [str(f.Name) for f in mia_pu.RowFields]
            check(page_value(mia_ho, "start_time_type") == "Handover Time",
                  "MIA Handover 表 start_time_type=Handover Time")
            check("Pickup Time" in page_visible(mia_pu, "start_time_type")
                  and "Handover Time" not in page_visible(mia_pu, "start_time_type"),
                  f"MIA Pick Up 表勾选 Pickup Time：{page_visible(mia_pu, 'start_time_type')}")
            mia_wb.Close(SaveChanges=False)
        wb.Close(SaveChanges=False)
        app.Quit()
        for prefix, label in (("Hub", "Hub Pivot"), ("OnTime", "On Time Rate")):
            cross = {k: v for k, v in totals.items() if k.startswith(prefix) and k.endswith("_cross")}
            entry = {k: v for k, v in totals.items() if k.startswith(prefix) and k.endswith("_entry")}
            elig = {k: v for k, v in totals.items() if k.startswith(prefix) and k.endswith("_elig")}
            check(set(cross.values()) == {"非跨Hub"},
                  f"{label} 四张表跨Hub转运默认 非跨Hub：{cross}")
            check(set(entry.values()) == {"(All)"},
                  f"{label} 四张表进表来源默认 (All)：{entry}")
            check(set(elig.values()).issubset({"True", "TRUE"}),
                  f"{label} 四张表 kpi_eligible 默认 True：{elig}")
        check(totals["HubHandoverPct_start"] == "Handover Time"
              and totals["HubHandoverQty_start"] == "Handover Time"
              and totals["OnTimeHandoverPct_start"] == "Handover Time"
              and totals["OnTimeHandoverQty_start"] == "Handover Time",
              "Handover 表 start_time_type=Handover Time")
        pickup_starts = [totals[f"{name}_start_items"] for name in
                         ("HubPickupPct", "HubPickupQty", "OnTimePickupPct", "OnTimePickupQty")]
        check(all("Pickup Time" in items and "Handover Time" not in items
                  and items <= {"Pickup Time", "Manifest Time"} for items in pickup_starts),
              f"Pick Up 表 start_time_type 勾选 Pickup Time / Manifest Time：{pickup_starts}")
        check(totals["Customer Pivot"] == "(All)",
              f"Customer Pivot 跨Hub转运默认 (All)：{totals['Customer Pivot']}")
        check(totals["Customer Pivot_entry"] == "(All)",
              f"Customer Pivot 进表来源默认 (All)：{totals['Customer Pivot_entry']}")
        check(str(totals["Customer Pivot_elig"]).upper() == "TRUE",
              f"Customer Pivot kpi_eligible 默认 True：{totals['Customer Pivot_elig']}")
        customer_row_keys = ["Customer Pivot_rows"] + [
            f"{name}_rows" for name in (
                "OnTimeHandoverPct", "OnTimeHandoverQty", "OnTimePickupPct", "OnTimePickupQty",
            )
        ]
        if "MIA_rows" in totals:
            customer_row_keys += ["MIA_rows", "MIA_pickup_rows"]
        check(all(totals.get(k) == ["CUSTOMER"] for k in customer_row_keys),
              "Customer 相关 Pivot 行字段为 CUSTOMER："
              + ", ".join(f"{k}={totals.get(k)}" for k in customer_row_keys))
    except Exception as exc:  # noqa: BLE001
        check(False, f"读取 Pivot 失败：{type(exc).__name__} {exc}")

    print("\n" + "=" * 60)
    passed = sum(1 for ok, _ in RESULTS if ok)
    print(f"自测结果：{passed}/{len(RESULTS)} 通过")
    for ok, label in RESULTS:
        if not ok:
            print(f"  FAIL {label}")
    print("=" * 60)
    return 0 if passed == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
