# -*- coding: utf-8 -*-
"""Native Excel PivotTables via local Excel automation. No static summary substitutes."""
from __future__ import annotations

import os
import shutil
import tempfile
import time
from pathlib import Path

import pandas as pd

from common import StopWithMessage, write_workbook

# XlConsolidationFunction
XL_COUNT = -4112
XL_AVERAGE = -4106
XL_SUM = -4157
# XlPivotFieldCalculation
XL_NO_CALC = -4143
XL_PERCENT_OF_ROW = 6
# XlPivotFieldOrientation
XL_ROW, XL_COLUMN, XL_PAGE = 1, 2, 3
XL_DATABASE = 1
XL_OPEN_XML_WORKBOOK = 51
XL_DO_NOT_REPEAT_LABELS = 2

PIVOT_SOURCE_SHEET = "Updated Unified Master"
MIA_SOURCE_SHEET = "MIA Source"
MIA_REPORT_NAME = "Ascan_KPI_MIA.xlsx"


def check_dependencies() -> tuple[bool, str]:
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, (
            "缺少 pywin32，无法创建 Excel 原生 Pivot。\n"
            "安装命令：python -m pip install pywin32"
        )
    try:
        import win32com.client as client
        app = client.DispatchEx("Excel.Application")
        version = app.Version
        app.Quit()
        return True, f"Excel {version} 可用"
    except Exception as exc:  # noqa: BLE001
        return False, f"本机 Excel 自动化不可用：{type(exc).__name__} {exc}"


def order_items(field, order: list[str]) -> None:
    """Force the requested display order; anything else keeps its natural position after."""
    present = {}
    for item in field.PivotItems():
        present[str(item.Name)] = item
    position = 1
    for name in order:
        item = present.get(name)
        if item is None:
            continue
        try:
            item.Position = position
            position += 1
        except Exception:  # noqa: BLE001 - Excel rejects moves for hidden items
            continue


def _page_wanted(page) -> list[str] | None:
    if page is None or page == "(All)":
        return None
    if isinstance(page, (list, tuple, set)):
        return [str(v) for v in page]
    return [str(page)]


def _match_pivot_item(names: dict, value: str):
    for candidate in (value, value.upper(), value.lower(), "TRUE", "True", "Yes"):
        item = names.get(candidate)
        if item is not None:
            return item
    return None


def _set_page_field(pt, field_name: str, page="(All)") -> None:
    field = pt.PivotFields(field_name)
    field.Orientation = XL_PAGE
    try:
        field.ClearAllFilters()
    except Exception:  # noqa: BLE001
        pass
    wanted = _page_wanted(page)
    if wanted is None:
        field.EnableMultiplePageItems = False
        field.CurrentPage = "(All)"
        return
    names = {str(item.Name): item for item in field.PivotItems()}
    matched = []
    seen = set()
    for value in wanted:
        item = _match_pivot_item(names, value)
        if item is None or id(item) in seen:
            continue
        matched.append(item)
        seen.add(id(item))
    if not matched:
        field.EnableMultiplePageItems = False
        field.CurrentPage = "(All)"
        return
    if len(matched) == 1:
        field.EnableMultiplePageItems = False
        field.CurrentPage = str(matched[0].Name)
        return
    field.EnableMultiplePageItems = True
    for item in matched:
        try:
            item.Visible = True
        except Exception:  # noqa: BLE001
            pass
    keep = {id(item) for item in matched}
    for item in names.values():
        if id(item) in keep:
            continue
        try:
            item.Visible = False
        except Exception:  # noqa: BLE001
            pass


def _set_page_all(pt, field_name: str) -> None:
    _set_page_field(pt, field_name, "(All)")


def _disable_subtotals(field) -> None:
    try:
        field.Subtotals = [False] * 12
    except Exception:  # noqa: BLE001
        try:
            field.Subtotals[1] = False
        except Exception:  # noqa: BLE001
            pass


def _ungroup_field(pt, field_name: str) -> None:
    """Undo Excel's automatic date grouping so the original datetime stays visible."""
    for _ in range(3):
        try:
            pt.PivotFields(field_name).DataRange.Ungroup()
            continue
        except Exception:  # noqa: BLE001
            pass
        try:
            pt.PivotFields(field_name).LabelRange.Ungroup()
            continue
        except Exception:  # noqa: BLE001
            return


def _apply_common_filters(pt, report: dict) -> None:
    _set_page_all(pt, "跨Hub转运")
    _set_page_all(pt, "进表来源")
    _set_page_field(pt, "kpi_eligible", report.get("kpi_eligible_filter_default", "True"))


def _apply_filters(pt, filters: dict | None, report: dict) -> None:
    if not filters:
        _apply_common_filters(pt, report)
        return
    for field_name, page in filters.items():
        _set_page_field(pt, field_name, page)


def _ensure_sheet(workbook, name: str):
    try:
        return workbook.Worksheets(name)
    except Exception:  # noqa: BLE001
        sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        sheet.Name = name
        return sheet


def _copy_to_dest(src: Path, dest: Path, logger, attempts: int = 8) -> Path:
    """Copy into a OneDrive/Excel-synced folder. Retry when the dest file is locked."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    lock = dest.with_name(f"~${dest.name}")
    tmp = dest.with_name(f"{dest.stem}.copying{dest.suffix}")
    last_err: OSError | None = None
    for i in range(attempts):
        try:
            if lock.exists():
                try:
                    lock.unlink()
                except OSError:
                    pass
            if tmp.exists():
                try:
                    tmp.unlink()
                except OSError:
                    pass
            shutil.copy2(src, tmp)
            os.replace(tmp, dest)
            return dest
        except OSError as exc:
            last_err = exc
            wait = min(30, 2 ** i)
            logger.warning("写回 %s 被占用，%d 秒后重试（%s）", dest.name, wait, exc)
            time.sleep(wait)
    fallback = dest.with_name(f"{dest.stem}_ready{dest.suffix}")
    shutil.copy2(src, fallback)
    logger.warning(
        "无法覆盖 %s（%s），已另存为 %s。请关闭占用该文件的程序后把此文件改回正式名。",
        dest.name, last_err, fallback.name,
    )
    return fallback


def _add_pivot(workbook, sheet, cache, spec: dict, report: dict, font: str, logger) -> None:
    dest_row = int(spec.get("dest_row", 3))
    dest_col = int(spec.get("dest_col", 1))
    title = spec.get("title")
    title_row = spec.get("title_row")
    dest = sheet.Cells(dest_row, dest_col)
    pt = cache.CreatePivotTable(TableDestination=dest, TableName=spec["name"])
    try:
        pt.ManualUpdate = True
    except Exception:  # noqa: BLE001
        pass
    _apply_filters(pt, spec.get("filters"), report)

    for i, row_name in enumerate(spec["rows"], start=1):
        field = pt.PivotFields(row_name)
        field.Orientation = XL_ROW
        field.Position = i
        if row_name in spec.get("subtotals_off", ()):
            _disable_subtotals(field)
        if row_name in spec.get("ungroup_dates", ()):
            _ungroup_field(pt, row_name)

    if spec.get("column"):
        pt.PivotFields(spec["column"]).Orientation = XL_COLUMN

    for val in spec["values"]:
        field = pt.AddDataField(pt.PivotFields(val["field"]), val["caption"], val["func"])
        field.Calculation = val.get("calc", XL_NO_CALC)
        field.NumberFormat = val["number_format"]
    if len(spec["values"]) > 1:
        try:
            pt.DataPivotField.Orientation = XL_COLUMN
        except Exception:  # noqa: BLE001
            pass

    if spec.get("column"):
        order_items(pt.PivotFields(spec["column"]),
                    spec.get("column_order") or report.get("status_order", []))
    if spec.get("row_order"):
        order_items(pt.PivotFields(spec["rows"][0]), spec["row_order"])
    for field_name, order in spec.get("field_orders", {}).items():
        order_items(pt.PivotFields(field_name), order)

    pt.RowAxisLayout(1)  # xlTabularRow
    try:
        pt.RepeatAllLabels(XL_DO_NOT_REPEAT_LABELS)
    except Exception:  # noqa: BLE001
        pass
    try:
        pt.ManualUpdate = False
    except Exception:  # noqa: BLE001
        pass
    for row_name in spec.get("ungroup_dates", ()):
        _ungroup_field(pt, row_name)
    for field in pt.PageFields:
        try:
            source = str(field.SourceName)
            if source and str(field.Name) != source:
                field.Name = source
        except Exception:  # noqa: BLE001
            pass
    try:
        pt.TableRange2.Font.Name = font
    except Exception:  # noqa: BLE001
        pass
    if title:
        try:
            top = int(pt.TableRange2.Row)
        except Exception:  # noqa: BLE001
            top = dest_row
        write_row = int(title_row) if title_row and int(title_row) < top else max(1, top - 1)
        if write_row < top:
            cell = sheet.Cells(write_row, dest_col)
            cell.Value = title
            cell.Font.Name = font
            cell.Font.Bold = True
    try:
        rng = pt.TableRange2
        if int(rng.Rows.Count) <= 1 and int(rng.Columns.Count) <= 1:
            logger.warning("Pivot %s / %s 只有 1 格，可能是 #SPILL!", sheet.Name, spec["name"])
    except Exception:  # noqa: BLE001
        pass
    logger.info("  已创建原生 Pivot：%s / %s", sheet.Name, spec["name"])


def _sheet_last_pivot_row(sheet) -> int:
    last = 0
    for pt in sheet.PivotTables():
        try:
            rng = pt.TableRange2
            last = max(last, int(rng.Row + rng.Rows.Count - 1))
        except Exception:  # noqa: BLE001
            continue
    return last


PICKUP_START_TYPES = ["Pickup Time", "Manifest Time"]


def _hub_split_filters(start_time_type, report: dict) -> dict:
    return {
        "kpi_eligible": report.get("kpi_eligible_filter_default", "True"),
        "进表来源": "(All)",
        "跨Hub转运": report.get("hub_pivot_cross_hub_default", "非跨Hub"),
        "start_time_type": start_time_type,
    }


def pivot_sheet_specs(report: dict, pct: str) -> list[dict]:
    hub_order = report.get("hub_order", [])
    count_fmt = report.get("count_format", "#,##0")
    handover_filters = _hub_split_filters("Handover Time", report)
    pickup_filters = _hub_split_filters(PICKUP_START_TYPES, report)
    hub_rows = {
        "rows": ["Hub"], "column": "status", "row_order": hub_order,
    }
    ontime_rows = {
        "rows": ["CUSTOMER"], "column": "Hub", "row_order": [],
        "column_order": hub_order,
    }
    return [
        {
            "sheet": "Hub Pivot",
            "pivots": [
                {
                    **hub_rows, "name": "HubHandoverPct", "title": "Handover 百分比",
                    "title_row": 1, "dest_row": 6, "dest_col": 1,
                    "filters": handover_filters,
                    "values": [{
                        "field": "TNO", "func": XL_COUNT, "calc": XL_PERCENT_OF_ROW,
                        "caption": "Count of TNO", "number_format": pct,
                    }],
                },
                {
                    **hub_rows, "name": "HubHandoverQty", "title": "Handover 数量",
                    "title_row": 1, "dest_row": 6, "dest_col": 13,
                    "filters": handover_filters,
                    "values": [{
                        "field": "TNO", "func": XL_COUNT, "calc": XL_NO_CALC,
                        "caption": "Count of TNO", "number_format": count_fmt,
                    }],
                },
                {
                    **hub_rows, "name": "HubPickupPct", "title": "Pick Up / Manifest 百分比",
                    "title_row": 45, "dest_row": 50, "dest_col": 1,
                    "filters": pickup_filters,
                    "values": [{
                        "field": "TNO", "func": XL_COUNT, "calc": XL_PERCENT_OF_ROW,
                        "caption": "Count of TNO", "number_format": pct,
                    }],
                },
                {
                    **hub_rows, "name": "HubPickupQty", "title": "Pick Up / Manifest 数量",
                    "title_row": 45, "dest_row": 50, "dest_col": 13,
                    "filters": pickup_filters,
                    "values": [{
                        "field": "TNO", "func": XL_COUNT, "calc": XL_NO_CALC,
                        "caption": "Count of TNO", "number_format": count_fmt,
                    }],
                },
            ],
        },
        {
            "sheet": "Customer Pivot",
            "pivots": [{
                "name": "CustomerPivot",
                "rows": ["CUSTOMER"], "column": "status",
                "dest_row": 3, "dest_col": 1,
                "values": [{
                    "field": "TNO", "func": XL_COUNT, "calc": XL_PERCENT_OF_ROW,
                    "caption": "Count of TNO", "number_format": pct,
                }],
                "row_order": [],
            }],
        },
        {
            "sheet": "On Time Rate",
            "pivots": [
                {
                    **ontime_rows, "name": "OnTimeHandoverPct", "title": "Handover 百分比",
                    "title_row": 1, "dest_row": 6, "dest_col": 1,
                    "filters": handover_filters,
                    "values": [{
                        "field": "onTime_count", "func": XL_AVERAGE, "calc": XL_NO_CALC,
                        "caption": "A-Scan On-Time Rate / A-Scan 上网及时率", "number_format": pct,
                    }],
                },
                {
                    **ontime_rows, "name": "OnTimeHandoverQty", "title": "Handover 数量",
                    "title_row": 1, "dest_row": 6, "dest_col": 16,
                    "filters": handover_filters,
                    "values": [{
                        "field": "TNO", "func": XL_COUNT, "calc": XL_NO_CALC,
                        "caption": "Count of TNO", "number_format": count_fmt,
                    }],
                },
                {
                    **ontime_rows, "name": "OnTimePickupPct", "title": "Pick Up / Manifest 百分比",
                    "title_row": 80, "dest_row": 85, "dest_col": 1,
                    "filters": pickup_filters,
                    "values": [{
                        "field": "onTime_count", "func": XL_AVERAGE, "calc": XL_NO_CALC,
                        "caption": "A-Scan On-Time Rate / A-Scan 上网及时率", "number_format": pct,
                    }],
                },
                {
                    **ontime_rows, "name": "OnTimePickupQty", "title": "Pick Up / Manifest 数量",
                    "title_row": 80, "dest_row": 85, "dest_col": 16,
                    "filters": pickup_filters,
                    "values": [{
                        "field": "TNO", "func": XL_COUNT, "calc": XL_NO_CALC,
                        "caption": "Count of TNO", "number_format": count_fmt,
                    }],
                },
            ],
        },
    ]


def mia_hub_sheet_specs(report: dict, pct: str) -> list[dict]:
    handover_filters = _hub_split_filters("Handover Time", report)
    pickup_filters = _hub_split_filters(PICKUP_START_TYPES, report)
    layout = {
        "rows": ["CUSTOMER"], "column": "Hub", "row_order": [],
        "column_order": report.get("hub_order", []),
    }
    return [{
        "sheet": "Pivot",
        "pivots": [
            {
                **layout, "name": "MIAHandoverRate", "title": "Handover",
                "title_row": 1, "dest_row": 6, "dest_col": 1,
                "filters": handover_filters,
                "values": [{
                    "field": "onTime_count", "func": XL_AVERAGE, "calc": XL_NO_CALC,
                    "caption": "A-Scan On-Time Rate / A-Scan 上网及时率", "number_format": pct,
                }],
            },
            {
                **layout, "name": "MIAPickupRate", "title": "Pick Up / Manifest",
                "title_row": 28, "dest_row": 33, "dest_col": 1,
                "filters": pickup_filters,
                "values": [{
                    "field": "onTime_count", "func": XL_AVERAGE, "calc": XL_NO_CALC,
                    "caption": "A-Scan On-Time Rate / A-Scan 上网及时率", "number_format": pct,
                }],
            },
        ],
    }]


def split_mia_hub_frames(report: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """MIA detail tabs plus the hidden pivot source. TTFBT is all hubs."""
    import kpi_rules
    from dc_bridge import merge_module

    empty = report.iloc[0:0].copy()
    hub = report["Hub"].astype("string").str.strip().str.upper() if "Hub" in report.columns else pd.Series(
        "", index=report.index, dtype="string"
    )
    mia = report.loc[hub.eq("MIA")].copy()
    start = kpi_rules.start_type_series(mia)
    handover = mia.loc[kpi_rules.is_handover_time_type(start)].copy()
    pickup = mia.loc[
        kpi_rules.is_pickup_time_type(start) | kpi_rules.is_manifest_time_type(start)
    ].copy()
    fbt = report.loc[merge_module().is_fbt_row(report)].copy()
    return {
        "Handover": handover if len(handover) else empty,
        "PickUp": pickup if len(pickup) else empty,
        "TTFBT": fbt if len(fbt) else empty,
        MIA_SOURCE_SHEET: mia if len(mia) else empty,
    }


def build_mia_hub_workbook(path: Path, report: pd.DataFrame, cfg: dict, logger) -> None:
    frames = split_mia_hub_frames(report)
    logger.info("MIA Hub 报表明细：Handover %d / PickUp %d / TTFBT %d / MIA 源 %d",
                len(frames["Handover"]), len(frames["PickUp"]),
                len(frames["TTFBT"]), len(frames[MIA_SOURCE_SHEET]))
    return build_workbook(
        path, frames, cfg, logger,
        source_sheet=MIA_SOURCE_SHEET,
        sheet_specs=mia_hub_sheet_specs(cfg["report"], cfg["report"].get("percent_format", "0.00%")),
        hide_sheets=(MIA_SOURCE_SHEET,),
        activate="Pivot",
    )


def build_workbook(path: Path, sheets: dict[str, pd.DataFrame],
                   cfg: dict, logger, *,
                   source_sheet: str = PIVOT_SOURCE_SHEET,
                   sheet_specs: list[dict] | None = None,
                   hide_sheets: tuple[str, ...] = (),
                   activate: str | None = None) -> None:
    """Write the data with xlsxwriter first (fast), then let Excel add real pivots."""
    import win32com.client as client

    report = cfg["report"]
    font = report.get("font_name", "微软雅黑")
    pct = report.get("percent_format", "0.00%")
    if source_sheet not in sheets:
        raise StopWithMessage(f"报表缺少 {source_sheet}，无法创建 Pivot。")
    pivot_source = sheets[source_sheet]

    # Write and pivot on a local temp file. OneDrive can lock or crash Excel
    # if it opens a just-written 100MB+ workbook in the synced folder.
    local_dir = Path(tempfile.gettempdir()) / "closed_loop_pivots"
    local_dir.mkdir(parents=True, exist_ok=True)
    local_path = local_dir / path.name
    write_workbook(local_path, sheets, font=font)
    path.parent.mkdir(parents=True, exist_ok=True)
    logger.info("数据表已写入本地临时文件，接下来用 Excel 创建原生 Pivot（数据源 %s）",
                source_sheet)

    app = client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    try:
        app.EnableEvents = False
        app.AutoRecover.Enabled = False
    except Exception:  # noqa: BLE001
        pass
    workbook = None
    saved_as = None
    try:
        logger.info("打开本地工作簿创建 Pivot：%s", local_path)
        workbook = app.Workbooks.Open(str(local_path), UpdateLinks=0, ReadOnly=False)
        src_ws = workbook.Worksheets(source_sheet)
        last_row = len(pivot_source) + 1
        last_col = max(len(pivot_source.columns), 1)
        src_range = src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(last_row, last_col))
        specs_to_build = sheet_specs if sheet_specs is not None else pivot_sheet_specs(report, pct)

        if pivot_source.empty:
            logger.warning("%s 无数据行，跳过 Pivot 创建", source_sheet)
            specs_to_build = []

        cache = (
            workbook.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=src_range)
            if specs_to_build else None
        )
        for sheet_spec in specs_to_build:
            sheet = _ensure_sheet(workbook, sheet_spec["sheet"])
            specs = [dict(spec) for spec in sheet_spec["pivots"]]
            prev_dest_row = 0
            for spec in specs:
                dest_row = int(spec.get("dest_row", 3))
                if dest_row > prev_dest_row > 0:
                    last = _sheet_last_pivot_row(sheet)
                    shift = last + 6 - dest_row
                    if shift > 0:
                        band = dest_row
                        for later in specs:
                            if int(later.get("dest_row", 3)) == band:
                                later["dest_row"] = int(later["dest_row"]) + shift
                                if later.get("title_row"):
                                    later["title_row"] = int(later["title_row"]) + shift
                        dest_row = int(spec["dest_row"])
                prev_dest_row = dest_row
                _add_pivot(workbook, sheet, cache, spec, report, font, logger)
            sheet.Columns("A:AZ").AutoFit()

        for name in sheets:
            try:
                ws = workbook.Worksheets(name[:31])
                ws.Cells.Font.Name = font
            except Exception:  # noqa: BLE001
                continue
        # Save before Activate. Excel/RPC can die on Activate after a long
        # pivot build; the file must already be on disk if that happens.
        try:
            workbook.Save()
        except Exception:  # noqa: BLE001
            saved_as = local_path.with_name(f"{local_path.stem}_pivot{local_path.suffix}")
            workbook.SaveAs(str(saved_as), FileFormat=XL_OPEN_XML_WORKBOOK)
        for hidden in hide_sheets:
            try:
                workbook.Worksheets(hidden[:31]).Visible = False
            except Exception:  # noqa: BLE001
                logger.warning("隐藏工作表失败，忽略：%s", hidden)
        try:
            first = activate or (sheet_specs[0]["sheet"] if sheet_specs else None)
            if first:
                workbook.Worksheets(first[:31]).Move(Before=workbook.Worksheets(1))
                workbook.Worksheets(first[:31]).Activate()
            else:
                workbook.Worksheets(1).Activate()
            workbook.Save()
        except Exception as activate_exc:  # noqa: BLE001
            logger.warning("Pivot 已保存，但 Activate 失败，忽略：%s", activate_exc)
    except Exception as exc:  # noqa: BLE001
        raise StopWithMessage(
            f"创建 Excel 原生 Pivot 失败：{type(exc).__name__} {exc}\n"
            f"数据表已写入 {path}，但没有 Pivot。请确认本机 Excel 可用后重跑第二步。"
        ) from exc
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001
            logger.warning("关闭 Excel 工作簿失败，忽略")
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            logger.warning("退出 Excel 失败，忽略")
    finished = saved_as if saved_as is not None else local_path
    written = _copy_to_dest(finished, path, logger)
    logger.info("已将带 Pivot 的工作簿写回 %s", written.name)
    try:
        if saved_as is not None and saved_as.exists():
            saved_as.unlink()
        if local_path.exists():
            local_path.unlink()
    except OSError:
        pass
    return written
