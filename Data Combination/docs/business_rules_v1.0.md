# 每日 A-Scan 数据整合业务规则 V1.0

配置源：`config/business_rules_v1.0.yaml`。原始 Excel 只读，不得修改、重命名或移动。

## 去重

- 主键：清洗后的 `TNO`（tracking number，文本，保留前导零）
- 一单一行
- 按字段 outer 合并，禁止 append 后再按单号删行
- `Source`、`File名称`、冲突审计保留全部来源

## 来源优先级

同一 TNO、同一字段：

1. Platform 有有效值 → Platform
2. Platform 为空 → System
3. System 为空 → AGS

`primary_source` 有 Platform 则为 `Platform`，否则 System，再否则 AGS。  
`Source` 与 `File名称` 用分号连接全部来源，不去重到只留一个。

## System Hub

System（`handover_pickup...`）的 Hub：

1. `first_ascan_hub` 有值 → 用它，**且优先于 Platform / AGS**，不被 Origin Hub 覆盖
2. `first_ascan_hub` 为空 → 用 System `origin_poe`；此时仍是 Platform 有值用 Platform，否则 System，再否则 AGS
3. System 两列都空 → **删除该 System 行**，不进入统一主表（该 TNO 仍可由 AGS / Platform 进入）

`first_ascan_poe` 仍是 A-scan **时间**，不是 Hub。其他字段（Customer、时间、MAWB）仍是 Platform → System → AGS。

DFW 没有 TT FBT。主表 Hub 选定之后，TT FBT（`CUSTOMER=TT FBT` / 业务名称含 FBT / `platform_sheet=FBT`）若 Hub 为 DFW，一律改成 IAH。不改 `跨Hub转运`。

## Customer 标准化

主表 `CUSTOMER` 为标准化名称；`customer_raw` 为各来源原始 Customer，去重后分号连接。

| 原始 / 来源 | CUSTOMER |
|---|---|
| AGS `TIKTOK` / `TikTok Inc` / 含 TikTok 的变体 | `TT GS` |
| `HANGZHOU CAINIAO SUPPLY CHAIN MANAGEMENT CO.,LTD.` | `AliExpress` |
| Platform AliExpress | `AliExpress` |
| Platform FBT | `TT FBT` |
| Platform Winit | `Winit` |
| Platform TT-GS | `TT GS` |

## start_time_type 展示

| 场景 | 展示 |
|---|---|
| Handover | `Handover Time` |
| Pickup | `Pickup Time` |
| FBT Manifest Time | `Manifest Time` |
| Winit 原始 `date_manifested` | `Manifest Time` |

Winit 起算时间仍取 `date_manifested` 的值，只改展示名称。

## SLA 与 status

- TT FBT 且 Hub=MIA：`Ascan Due Date = start_time + 12 小时`
- 其他 Hub 的 TT FBT：`start_time + 24 小时`
- 仅 System（Source 不含 AGS）的 ProCarrier、Vinted US Domestic：`start_time + 48 小时`
- Source 含 AGS 的同一客户仍为 24 小时
- **Winit 常规：日期制 D1**（Manifested Time 当地日期 + 1 个自然日，截止小于次日 00:00:00；展示 23:59:59）。不跳过周末。
- **Winit `Hub` 含「自提」：日期制 D0**（出库当天结束前）。不要把所有 Winit LAX 当成自提。
- 其他小时制业务：`start_time + 24 小时`
- `cohort_date` = `start_time` 本地日期（出库/交接日，不是 KPI 归属日）
- `time_diff_hours` = `first_ascan_time - start_time`（分析用；Winit 不以它判断达标）
- Winit 客户目标 95%；内部目标 99.5%，分别标识

| status | 条件 |
|---|---|
| On Time | 有 A-Scan 且 `first_ascan_time <= Ascan Due Date` |
| Late | 有 A-Scan 且 `first_ascan_time > Ascan Due Date` |
| No A-Scan | `first_ascan_time` 为空 |

不输出 `Pending`。未到截止时间且无 A-Scan 也记为 `No A-Scan`。

## 主表 18 列顺序

1. Source  
2. File名称  
3. primary_source  
4. 业务名称  
5. TNO  
6. MAWB  
7. CUSTOMER  
8. customer_raw  
9. Hub  
10. start_time  
11. start_time_type  
12. first_ascan_time  
13. Ascan Due Date  
14. cohort_date  
15. time_diff_hours  
16. sla_hours（小时制 SLA；Winit 为空）
17. sla_type（hours / calendar_d1 / calendar_d0）
18. status  
19. 跨Hub转运  

前 17 列除 Hub 外仍按 Platform → System → AGS。Hub 见上文 System Hub：`first_ascan_hub` 优先于其他来源。`跨Hub转运` 是单独的 AGS 转运风险标签，不参与该字段优先级。

## 跨Hub转运

只允许 `LAX-DFW` / `EWR-ORD`，其余空白。

识别只用 AGS 原始字段，再按 TNO 回写主表：

- `ags_mawb_hub` = Tracking Details.mawb 匹配 MAWB Details.mawb 后的 `MAWB Details.Hub`
- `ags_destination_hub` = Tracking Details.destination_hub
- `Tracking Details.origin_hub` 仅审计，不作为主判断
- **禁止**用 unified_master 最终 `Hub` 判断（该列可能是 Platform Origin POE）

按 `ags_mawb_hub + mawb` 汇总 destination_hub 去重集合：

- LAX 且集合 ⊆ `{DFW, IAH}` → 该 MAWB 全部 TNO = `LAX-DFW`
- EWR 且集合同时包含 `{ATL, EWR, MIA, ORD}` → 该 MAWB 全部 TNO = `EWR-ORD`

MAWB Details 同一 MAWB 对应多个 Hub 时，`ags_mawb_hub` 留空，写入 conflict_audit，不自动打标。

主表不删除跨 Hub 包裹。当前正式 A-Scan Rate 暂排除 `跨Hub转运` 非空行。明细见 `output/inter_hub_transfer_detail_YYYY-MM-DD.xlsx`。

后续有实际接收时间后，保留原始 start_time / Ascan Due Date，写入 `actual_receiving_time` 再重算 SLA；完成回填后才重新纳入最终 Rate。
