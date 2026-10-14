# Changelog

## 1.1.4 — 2026-09-21

DFW 没有 TT FBT。主表 Hub 选定后，TT FBT 且 Hub=DFW 的记录改成 IAH。闭环第一步 / 第二步同样改写，Hub Pivot、MAWB 汇总、Pool 都按 IAH。`跨Hub转运` 标签不变。Platform TT-GS 明细 Hub 列增加别名 `实际入网口岸`。

## 1.1.3 — 2026-09-15

统一主表 Hub：System `first_ascan_hub` 优先于 Platform / AGS，不再被 Origin Hub 覆盖。没有 `first_ascan_hub` 时仍按 Platform → System `origin_poe` → AGS。Platform 明细表别名增加 `cainiao 明细`（AliExpress）。

## 1.1.2 — 2026-09-15

System Hub 优先取 `first_ascan_hub`，为空再用 `origin_poe`。两列都空的 System 行不进入统一主表。`first_ascan_poe` 仍是 A-scan 时间，不是 Hub。Platform 明细表别名增加 `Parcels Detail`（TT-GS）和 `明细`（AliExpress）。

## 1.1.1 — 2026-09-07

仅 System（不含 AGS）的 ProCarrier、Vinted US Domestic 使用 `start_time + 48 小时`。同一客户只要 Source 含 AGS，仍为 24 小时。TT FBT 仍为 12 小时，其余仍为 24 小时。

## 1.1.0 — 2026-08-31

新增第 18 列 `跨Hub转运`。仅用 AGS 的 MAWB Hub 与 destination_hub 识别，按 TNO 回写；不使用主表最终 Hub。允许值只有 `LAX-DFW` / `EWR-ORD`。跨 Hub 包裹保留在主表，但暂不纳入正式 A-Scan Rate。输出 `inter_hub_transfer_detail_YYYY-MM-DD.xlsx`。

## 1.0.0 — 2026-08-31

每日 A-Scan 数据整合基线。已按本批三源数据核对：一单一行、Platform 优先、Customer 标准化、Manifest Time 展示、12H/24H SLA。

### 去重

- 主键 TNO，一单一行
- 字段级 outer 合并，禁止 append 后删重复行
- Source / File名称保留全部来源

### 字段优先级

Platform → System → AGS；空值不覆盖下一来源的有效值。

### Customer 标准化

- AGS TikTok 变体 → `TT GS`
- Cainiao 全称 → `AliExpress`
- Platform FBT / Winit / AliExpress / TT-GS 分别为 `TT FBT` / `Winit` / `AliExpress` / `TT GS`
- `customer_raw` 保留各来源原始值（分号连接）
- 主表展示 `CUSTOMER`

### 起算时间类型

- Pickup / Handover 分别显示 `Pickup Time` / `Handover Time`
- FBT 与 Winit 主表均显示 `Manifest Time`（Winit 仍用 `date_manifested` 时间起算）

### SLA 与 status

- TT FBT：`start_time + 12h`；其他：`+ 24h`
- status 仅 `On Time` / `Late` / `No A-Scan`，无 Pending

### 主表结构

对外主表固定 17 列，见 `docs/business_rules_v1.0.md`。
