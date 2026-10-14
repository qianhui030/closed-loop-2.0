# 每日 A-Scan 数据整合

**当前稳定版本：V1.1.3**

将同一业务日的 AGS、System、Platform 三份 Excel 按 TNO 合并为一单一行主表，不做原始文件改写。

## 运行

把当日三份源 Excel 放在本目录后执行：

```bash
python build_unified_master.py
```

仅用上次合并结果重导主表（不重读原始 Excel）：

```bash
python build_unified_master.py --from-csv
```

依赖：`pandas`、`openpyxl`、`xlsxwriter`、`pyarrow`。见 `requirements.txt`。

## 输出

主文件：`output/unified_master_YYYY-MM-DD.xlsx`（另有 `.csv` / `.parquet`）

18 列：前 17 列不变，第 18 列为 `跨Hub转运`（仅 `LAX-DFW` / `EWR-ORD`，其余空白）。该列只用 AGS 原始 MAWB Hub 与 destination_hub 识别，再按 TNO 回写，不用主表最终 Hub。

跨 Hub 明细：`output/inter_hub_transfer_detail_YYYY-MM-DD.xlsx`。当前正式 A-Scan Rate 暂排除该列非空行。

规则说明：`docs/business_rules_v1.0.md`  
机器可读配置：`config/business_rules_v1.0.yaml`
