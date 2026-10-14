# 每日三源数据：结构分析与自动化方案设计

**分析日期：** 2026-08-30  
**分析范围：** 当前文件夹中的三份原始 Excel，只读、不修改。  
**本阶段结论：** 三份文件可以合并为一单一行主数据；**不要** append 后按单号删重复。字段冲突时按来源优先级取值，并保留来源追溯和冲突审计。

---

## 0. 业务日期识别结果

**识别出的当天业务日期：`2026-08-24`。**

识别依据（按优先级）：

| 依据 | 结果 |
|---|---|
| AGS 文件名 | `AGS_to_SPX_24H_Ascan_Report_2026-08-24.xlsx` → 8/24 |
| System 文件名 | `handover_pickup_20260824.xlsx` → 8/24 |
| Platform 文件名 | `8.26平台客户需ASCAN单号明细.xlsx` → **8.26，与另外两份不一致** |
| AGS `handoverTime_local` | 151,995 / 151,995 均为 **2026-08-24** |
| System `pickup_poe_51005` | 213,510 条，全部 **2026-08-24** |
| System `handover_time_local` | 33,203 条，全部 **2026-08-24** |
| Platform TT-GS `交接当地时间` | 35,346 / 35,346 均为 **2026-08-24** |
| Platform AliExpress `清关交付-SPX收货时间` | 18,211 / 18,211 均为 **2026-08-24** |

补充说明：

- AGS 工作表 `24H Ascan Rate View` 的 `Date` 为 **8/26/2026**，与文件名 8/24 不同。结合 24H A-Scan 窗口（交接日 + 次日），该日期更像是**报表查看日/导出日**，不是 cohort 日。
- Platform 文件名中的 `8.26` 与上述查看日一致，**不能**当作 cohort 日。
- AGS 的 `handoverTime_UTC` 有 50,350 条落到 8/25（本地仍是 8/24）。**业务日期必须用本地交接/揽收日期，不能用 UTC，也不能用 A-Scan 日期。**
- A-Scan 日期分布跨 8/22–8/26，符合“后续节点不影响当天 cohort”的口径。

**三个文件并不完全是同一业务日：**

- AGS、System、Platform 的 GS / AliExpress / 大部分 Winit：**8/24 主 cohort**。
- Platform **FBT 明细** 24,512 / 24,994（98.1%）`Manifest Time` 为 **2026-08-25**。
- Platform **Winit 明细** 701 条 `出库日期` / `date_manifested` 为 **2026-08-25**。

是否把 FBT/Winit 的 8/25 记录纳入 8/24 分母，需要业务确认（见文末待确认事项）。本报告统计时**未删除**这些记录，仅单独标记。

---

## 1. 文件和工作表结构

### 1.1 文件总览

| 推测类型 | 文件名 | 格式 | 大小 | 判断依据 |
|---|---|---|---|---|
| **AGS** | `AGS_to_SPX_24H_Ascan_Report_2026-08-24.xlsx` | `.xlsx` | 7.75 MB | 文件名含 `AGS_to_SPX`；仅有 Handover，无 Pickup；客户来自 MAWB 维表；含 24H A-Scan 汇总 |
| **System** | `handover_pickup_20260824.xlsx` | `.xlsx` | 10.15 MB | 文件名含 `handover_pickup`；同时有 Pickup 和 Handover 列；客户字段直接在明细中 |
| **Platform** | `8.26平台客户需ASCAN单号明细.xlsx` | `.xlsx` | 13.57 MB | 多平台工作表（TikTok GS / FBT / Winit / AliExpress）；含透视汇总与未上网异常表 |

---

### 1.2 文件 1 — AGS

**主明细表：`Tracking Details`**（151,995 数据行 × 11 列）。  
理由：唯一达到包裹粒度的工作表；其余为 MAWB 维表或 24H 汇总。

| 工作表 | 总行数（含表头） | 总列数 | 角色 |
|---|---|---|---|
| 24H Ascan Rate View | 10 | 11 | 按客户 × Hub 的 24H 率，查看日 8/26 |
| MAWB Details | 106 | 9 | MAWB 级客户/Hub/ATA/交接/票量 |
| **Tracking Details** | **151,996** | **11** | **主明细** |
| 24H Ascan Rate | 10 | 8 | 客户 × Hub 率（无日期列） |

**Tracking Details 字段：**  
`mawb`, `tracking_number`, `origin_hub`, `destination_hub`, `destination_facility`, `handoverTime_UTC`, `handoverTime_origin_hub_local_time`, `handoverTime_local`, `ascanTime`, `time_diff_hours`, `within_24h`

**无客户名称列。** 客户必须通过 `mawb` 关联 `MAWB Details.Customer`。本批 151,995 行全部匹配成功。

**MAWB Details 字段：**  
`MAWB`, `Customer`, `Hub`, `ATA_time_UTC`, `ATA_time_local`, `handoverTime_UTC`, `handoverTime_local`, `volume`, `rate_24h_percent`

AGS 客户（按包裹量）：SHEIN Distribution Corporation 86,230；Whaleco Services, LLC 20,392；HANGZHOU CAINIAO SUPPLY CHAIN MANAGEMENT CO.,LTD. 14,584；4PX DIGITAL PARCEL EXPRESS CO., LIMITED 13,933；**TIKTOK 12,468**；Pro Carrier T11 1,791；Albatross America INC 1,748；cupshe 849。

**脱敏样例（Tracking Details）：**

| mawb | tracking_number | origin_hub | dest_hub | dest_facility | handoverTime_local | ascanTime | within_24h |
|---|---|---|---|---|---|---|---|
| 369-10881205 | SPXMIA****1141 | MIA | MIA | MIA | 2026-08-24 10:08:36 | 2026-08-24 14:48:37 | true |

---

### 1.3 文件 2 — System

**主明细表：`Sheet1`**（246,713 数据行 × 8 列）。仅此一张表。

**字段：**  
`tracking_number`, `mawb`, `customer_name`, `origin_poe`, `first_ascan_poe`, `pickup_poe_51005`, `handover_time`, `handover_time_local`

列名与真实含义不一致，必须以数据为准：

| 原始列名 | 实际内容 | 填充率 |
|---|---|---|
| `first_ascan_poe` | **First A-Scan 时间**（datetime，不是 POE 代码） | 95.3% |
| `pickup_poe_51005` | **Pickup Time**（datetime；`51005` 像事件码） | 86.5% |
| `origin_poe` | Origin Hub 代码（LAX/EWR/ORD 等） | 99.98% |
| `handover_time` | Handover UTC 或系统时区时间 | 13.5% |
| `handover_time_local` | Handover 本地时间 | 13.5% |

客户构成：Whaleco Services, LLC 153,713；Vinted US Domestic 37,308；**TikTok Inc 33,203**；SHEIN 18,097；ProCarrier 3,854；Fulstar America Corporation 386；GT ACCESSORY TRADING COMPANY LTD 93；REALTIME EXPRESS INC 59。

**关键结构事实：** Pickup 与 Handover 在本文件中**互斥**，从未出现在同一行、也从未出现在同一单号上。

- Pickup 非空 213,510：全部是非 TikTok（本地仓揽收）
- Handover 非空 33,203：全部是 TikTok Inc
- 两者同时非空：**0**

**脱敏样例：**

| tracking_number | customer_name | origin_poe | pickup_poe_51005 | handover_time_local | first_ascan_poe |
|---|---|---|---|---|---|
| SPXCLE****0881 | TikTok Inc | ORD | （空） | 2026-08-24 09:44:26 | 2026-08-24 17:40:13 |
| SPXMSP****9427 | TikTok Inc | ORD | （空） | 2026-08-24 09:44:26 | 2026-08-24 17:30:01 |

Pickup 非空样例时间形态：`2026-08-24 10:01:24`（datetime，无时区标记）。

---

### 1.4 文件 3 — Platform

多工作簿，**没有单一主表**，而是 4 张互斥的平台明细 + 透视汇总 + 未上网宽表。

四张明细之间 **0 个交叉单号**，并集 98,458 个唯一单号。

| 工作表 | 行数（含表头） | 列数 | 角色 | 是否主明细 |
|---|---|---|---|---|
| TT-GS透视 | 12 | 6 | 透视 | 否 |
| **TT-GS明细** | **35,347** | **18** | TikTok GS 包裹明细 | **是** |
| GS未上网明细 | 58 | 83 | GS 未上网异常宽表（57 单，全部已在 TT-GS明细中） | 否，子集 |
| GS PV | 13 | 3 | 透视 | 否 |
| TT-FBT透视 | 8 | 6 | 透视 | 否 |
| **FBT明细** | **24,995** | **10** | FBT 明细 | **是** |
| winit汇总 | 10 | 6 | 透视 | 否 |
| **winit明细** | **19,908** | **17** | Winit 明细 | **是** |
| AliExperess-汇总 | 11 | 6 | 透视 | 否 |
| **AliExperess明细** | **18,212** | **13** | AliExpress 明细 | **是** |
| FBT未上网明细 | 188 | 83 | FBT 未上网异常宽表（187 单，全部已在 FBT明细中） | 否，子集 |
| FBT PV | 9 | 5 | 透视 | 否 |

#### TT-GS明细（35,346 行）

字段：`追踪号`, `主运单号`, `邮袋号`, `卡转线路`, `州`, `分拣中心`, `邮编前三位`, `客户`, `交接当地时间`, `首次扫描当地时间`, `Ascan应履约日期`, `Ascan应履约日期取整`, `Ascan KPI`, `Origin POE`, `Destination POE`, `妥投当地时间`, `妥投天数`, `Ascan`

客户全部为 `TikTok Inc`。Ascan：准时上网 35,255；未上网 57；超时上网 34。

脱敏样例：`SPXABE****2972`，客户 TikTok Inc，交接当地时间 `2026-08-24 22:05:35`，Ascan=未上网。

#### FBT明细（24,994 行）

字段：`Tracking Number`, `Manifest Time`, `Manifest Date`, `First Arrival Scan`, `应履约日期`, `应履约日期取整`, `Ascan KPI`, `Origin Hub`, `Dest Hub`, `Ascan`

**无客户、无 MAWB、无 Handover。** `Manifest Time` 更像仓出库/揽收时间，且 98% 为 8/25。与 AGS、System **零重叠**。

#### winit明细（19,907 行）

字段：`Tracking Number`, `Origin Hub`, `Destination Hub`, `Destination Facility`, `Customer Code`, `Label Creation Date`, `First Scan Date`, `date_manifested`, `出库日期`, `ASCAN日期`, `ASCAN应履约日期`, `M-A`, `是否满足时效`, `满足时效`, `超时上网`, `未上网`, `Ascan`

Customer Code：WINIT_ORD 12,880；WINIT_LAX 3,664；WINIT 2,662；WINIT_LAX_COI 450；WINIT_LAX_WAT 251。与 AGS、System **零重叠**。

`Label Creation Date` / `First Scan Date` / `date_manifested` 为**文本**，带 `.0` 后缀，例如 `2026-08-24 18:48:27.0`。

#### AliExperess明细（18,211 行）

字段：`州`, `类型`, `跟踪号`, `大包号`, `提单号`, `清关交付-SPX收货时间`, `First Scan Arrival Scan`, `应履约日期`, `应履约日期取整`, `Ascan KPI`, `Origin POE`, `Destination POE`, `Ascan`

**无客户名称。** 16,324 单与 AGS 重叠，AGS 侧客户多为 Cainiao。`清关交付-SPX收货时间` 常见整分（如 `15:10:00`），精度低于 AGS。

#### 未上网宽表（补充字段来源，不计入主计数）

`GS未上网明细` / `FBT未上网明细` 含 Pickup Date、ZIP、重量、Last Physical Event、MAWB 等 83 列，但只覆盖未上网子集。后续可用于异常分析，**不能**替代主明细。

---

## 2. 标准字段映射分析

### 2.1 映射建议表

| 标准字段 | AGS | System | Platform-GS | Platform-FBT | Platform-Winit | Platform-Ali |
|---|---|---|---|---|---|---|
| tracking number | `tracking_number` | `tracking_number` | `追踪号` | `Tracking Number` | `Tracking Number` | `跟踪号` |
| customer name | `MAWB Details.Customer`（需 join） | `customer_name` | `客户` | **缺失** | `Customer Code`（仓代码，非客户全称） | **缺失** |
| platform | **缺失**（可从客户推断） | **缺失**（可从客户推断） | 工作表=TikTok-GS | 工作表=TikTok-FBT（待确认） | 工作表=Winit | 工作表=AliExpress；另有 `类型` |
| MAWB | `mawb` | `mawb`（仅 Handover 行基本有值） | `主运单号` | **缺失** | **缺失** | `提单号` |
| Pickup Time | **缺失** | `pickup_poe_51005` | **缺失** | `Manifest Time`（语义待确认） | `date_manifested`；另有日期粒度 `出库日期` | **缺失** |
| Handover Time | `handoverTime_local`（建议主用）；另有 UTC、origin hub local | `handover_time_local`；另有 `handover_time` | `交接当地时间` | **缺失** | `出库日期` 是否等于 Handover **待确认** | `清关交付-SPX收货时间` |
| First A-Scan Time | `ascanTime` | `first_ascan_poe`（列名误导） | `首次扫描当地时间` | `First Arrival Scan` | `First Scan Date` | `First Scan Arrival Scan` |
| Origin Hub | `origin_hub` | `origin_poe` | `Origin POE` | `Origin Hub` | `Origin Hub` | `Origin POE` |
| Destination Hub | `destination_hub` | **缺失** | `Destination POE` | `Dest Hub` | `Destination Hub` | `Destination POE` |
| Destination Facility | `destination_facility` | **缺失** | `分拣中心` | **缺失** | `Destination Facility` | **缺失** |
| Event Code | **缺失** | 仅列名中的 `51005` | **缺失** | **缺失** | **缺失** | **缺失** |
| Event Description | **缺失** | **缺失** | `Ascan`（准时/超时/未上网） | `Ascan` | `Ascan` | `Ascan` |
| Status | `within_24h`；`time_diff_hours` | **缺失** | `Ascan` + `Ascan KPI` | 同左 | `Ascan` + `是否满足时效` | `Ascan` + `Ascan KPI` |
| Address / ZIP | **缺失** | **缺失** | `州`, `邮编前三位` | **缺失** | **缺失** | `州` |

### 2.2 对后续 A-Scan / 异常分析有价值的其他字段

- AGS：`time_diff_hours`, `within_24h`；MAWB 级 `ATA_time_local` / `ATA_time_UTC`
- Platform-GS：`邮袋号`, `卡转线路`, `Ascan应履约日期`, `妥投当地时间`
- Platform-Winit：`Label Creation Date`, `M-A`
- Platform-Ali：`大包号`
- 未上网宽表：`Pickup Date`, `From Zip5`, `Last Physical Event*`, `Weight`

### 2.3 缺失、格式不一致与解析风险

**缺失（按来源）：**

- AGS 无 Pickup、无客户列（需 join）、无 platform、无地址。
- System 无 Destination Hub / Facility、无 Event Description、无地址；非 TikTok 行无 Handover；TikTok 行无 Pickup。
- Platform 各表字段不统一，不能当成一张表直接 union。

**日期格式：**

| 来源 | 列 | 存储类型 | 时区 | 解析风险 |
|---|---|---|---|---|
| AGS | handover / ascan | Excel datetime | 列名区分 UTC / local，单元格**无 tzinfo** | UTC 与 local 有 50,350 条跨日；必须明确用哪一列做 cohort |
| System | pickup / handover / ascan | Excel datetime | 无 tzinfo；handover UTC 与 local **日历日一致**（本批） | 列名 `first_ascan_poe` / `pickup_poe_51005` 易被误当成 Hub 代码 |
| Platform-GS | 交接/扫描时间 | **文本** `YYYY-MM-DD HH:MM:SS` | 无 | 需统一 parse；空扫描为空白字符串 |
| Platform-FBT | Manifest Time | datetime | 无 | 与 A-Scan 文本列混用 |
| Platform-Winit | date_manifested / First Scan / Label | **文本 + `.0`** | 无 | `2026-08-24 18:48:27.0` 必须去尾再解析 |
| Platform-Ali | 清关交付 | datetime，常见整分 | 无 | 精度低于 AGS/System，不宜直接当精确秒级冲突 |

**Tracking number 质量：**

- 三份文件主明细：**全部为字符串，无空值，无前后空格，无 float，无前导零。**
- 本批**没有**被 Excel 转成科学计数法的证据。
- 扫描到的 “E+数字” 命中是 **SPX 单号中字母 `E` 的误报**（如 `SPXABE006...`），不是科学计数法。
- 仍建议正式程序：一律按文本读入（`dtype=str`），`strip()`，拒绝 float 化。

**客户名称未对齐（同一实体多种写法）：**

- `TIKTOK`（AGS） vs `TikTok Inc`（System / Platform-GS）
- `SHEIN Distribution Corporation`（AGS） vs `SHEIN`（System）
- `Pro Carrier T11`（AGS） vs `ProCarrier`（System）

这不是不同客户，自动化必须先做客户别名标准化，否则会 100% 报“客户冲突”。

---

## 3. 单号唯一性和记录粒度

本批所有主明细都已经是 **一单一行**。

| 来源 | 总行数 | TN 非空 | 唯一 TN | 重复 TN 数 | 空 TN | 同一 TN 多行 |
|---|---|---|---|---|---|---|
| AGS Tracking Details | 151,995 | 151,995 | 151,995 | 0 | 0 | 否 |
| System Sheet1 | 246,713 | 246,713 | 246,713 | 0 | 0 | 否 |
| Platform-GS | 35,346 | 35,346 | 35,346 | 0 | 0 | 否 |
| Platform-FBT | 24,994 | 24,994 | 24,994 | 0 | 0 | 否 |
| Platform-Winit | 19,907 | 19,907 | 19,907 | 0 | 0 | 否 |
| Platform-Ali | 18,211 | 18,211 | 18,211 | 0 | 0 | 否 |
| Platform 四表明细并集 | 98,458 | 98,458 | 98,458 | 0（跨表也不交叉） | 0 | 否 |

### Pickup 与 Handover 是否同单共存

| 来源 | 结论 |
|---|---|
| AGS | 只有 Handover，无 Pickup 列 |
| System | **同一单号不会同时有 Pickup 和 Handover**；按客户类型拆开：本地仓=仅 Pickup，TikTok=仅 Handover |
| Platform-GS | 只有交接时间 |
| Platform-FBT | 只有 Manifest Time（更像 Pickup） |
| Platform-Ali | 只有清关交付时间 |

因此：用户说的“同一单号可能同时有 Pickup 和 Handover，两字段都要保留”在**本批真实数据中尚未出现**。Schema 仍应同时保留两列，但本批无需做“多事件折行聚合”。若以后同一单号在 System 出现两行（一次 Pickup、一次 Handover），才需要按 TN 聚合成一行、两列都保留。

---

## 4. 每日 cohort 日期检查

**业务日：2026-08-24。** A-Scan 日期不参与所属日判断。

### 4.1 Pickup Time 日期分布

| 来源 | 非空 | 2026-08-24 | 其他日 | 空 |
|---|---|---|---|---|
| AGS | 无此字段 | — | — | — |
| System `pickup_poe_51005` | 213,510 | **213,510** | 0 | 33,203（全部 TikTok Handover 行） |
| Platform-FBT `Manifest Time` | 24,994 | 482 | **8/25: 24,512** | 0 |
| Platform-Winit `date_manifested` | 19,907 | 19,206 | **8/25: 701** | 0 |

### 4.2 Handover Time 日期分布（建议用本地时间）

| 来源 | 非空 | 2026-08-24 | 其他日 | 空 |
|---|---|---|---|---|
| AGS `handoverTime_local` | 151,995 | **151,995** | 0 | 0 |
| AGS `handoverTime_UTC` | 151,995 | 101,645 | **8/25: 50,350** | 0 |
| System `handover_time_local` | 33,203 | **33,203** | 0 | 213,510（全部 Pickup 行） |
| Platform-GS `交接当地时间` | 35,346 | **35,346** | 0 | 0 |
| Platform-Ali `清关交付-SPX收货时间` | 18,211 | **18,211** | 0 | 0 |
| Platform-Winit `出库日期` | 19,907 | 19,206 | **8/25: 701** | 0 |

### 4.3 是否属于当天

判定规则：Pickup 本地日 = 8/24 **或** Handover 本地日 = 8/24。

| 来源 | 属于 8/24 | Pickup/Handover 均非当天 | 两者皆空 |
|---|---|---|---|
| AGS | 151,995（100%） | 0 | 0 |
| System | 246,713（100%） | 0 | 0 |
| Platform-GS | 35,346（100%） | 0 | 0 |
| Platform-Ali | 18,211（100%） | 0 | 0 |
| Platform-Winit | 19,206 | **701（8/25）** | 0 |
| Platform-FBT | 482 | **24,512（Manifest=8/25）** | 0 |

**未删除。** 8/25 样例（脱敏）：

- FBT：`SPXMFE****1416`，Manifest Time `2026-08-25 08:22:04`，First Arrival Scan `2026-08-23 18:00:42`（A-Scan 甚至早于 Manifest，进一步说明不能用 A-Scan 定 cohort，且 FBT 的 Manifest 语义需确认）
- Winit：`SPXCLL****0017`，出库日期 `2026-08-25`，date_manifested `2026-08-25 18:24:04.0`

System 没有“两者皆空”；AGS 没有 Pickup，但 Handover 全满。

---

## 5. 三个来源的单号重叠分析

清洗规则：去空白、按文本保留。本批无空单号、无科学计数法损坏。

### 5.1 三源 Venn

| 指标 | 单量 |
|---|---|
| AGS 有效单号 | 151,995 |
| System 有效单号 | 246,713 |
| Platform 有效单号（四表明细并集） | 98,458 |
| AGS ∩ System | 11,612 |
| AGS ∩ Platform | 27,936 |
| System ∩ Platform | 33,203 |
| 三份同时存在 | **11,612** |
| AGS 独有 | 124,059 |
| System 独有 | 213,510 |
| Platform 独有 | 48,931 |
| 三源并集（估算） | **436,027** |

要点：

- **AGS ∩ System 的 11,612 单全部也在 Platform-GS 中**，即三方重叠 = AGS ∩ System。
- System ∩ Platform 的 33,203 单 **等于 System 全部 TikTok**，且全部落在 Platform-GS。
- Platform-FBT、Winit 与 AGS/System **零重叠**，属于 Platform 独有。
- Platform-Ali 与 AGS 重叠 16,324，与 System **零重叠**。

按 Platform 子表拆开：

| Platform 子表 | 唯一单号 | ∩ AGS | ∩ System |
|---|---|---|---|
| TT-GS明细 | 35,346 | 11,612 | 33,203 |
| FBT明细 | 24,994 | 0 | 0 |
| winit明细 | 19,907 | 0 | 0 |
| AliExperess明细 | 18,211 | 16,324 | 0 |

### 5.2 TikTok 专项

识别方式：

- System：`customer_name` 含 TikTok → **TikTok Inc，33,203**
- AGS：join 后 `Customer` 含 TikTok → **TIKTOK，12,468**
- Platform-GS：客户全部 TikTok Inc，35,346
- Platform-FBT：无客户字段；与 System/AGS 零重叠，是否算 TikTok 需确认

| 指标 | 单量 |
|---|---|
| System 中 TikTok | 33,203 |
| AGS 中 TikTok（客户名=TIKTOK） | 12,468 |
| 两者重叠 | **11,612** |
| System TikTok 独有（不在 AGS） | **21,591** |
| AGS TikTok 独有（不在 System） | **856** |
| System TikTok ∩ Platform-GS | 33,203（System TikTok 全覆盖） |
| Platform-GS 比 System TikTok 多出 | 2,143 |
| AGS TikTok 不在 System | 856（这些单在 Platform-GS 中，属于“仅 AGS+Platform”的 TikTok） |

**重叠 TikTok 字段完整度（11,612 单，AGS vs System）：**

| 字段 | AGS 有值 | System 有值 | 双方都有且一致 | 冲突 | 仅 AGS | 仅 System |
|---|---|---|---|---|---|---|
| customer | 11,612 | 11,612 | 0 | **11,612**（`TIKTOK` vs `TikTok Inc`） | 0 | 0 |
| MAWB | 11,612 | 11,606 | 11,606 | 0 | 6 | 0 |
| Pickup | 0 | 0 | — | — | — | — |
| Handover | 11,612 | 11,612 | 9,346 | **2,266（19.5%，同日不同时刻）** | 0 | 0 |
| A-Scan | 11,578 | 11,578 | 11,578 | 0 | 0 | 0 |
| Origin Hub | 11,612 | 11,612 | 11,612 | 0 | 0 | 0 |
| Dest Hub | 11,612 | 0 | — | — | 11,612 | 0 |
| Dest Facility | 11,612 | 0 | — | — | 11,612 | 0 |

结论：TikTok 重叠单上，**System 并不比 AGS 更完整**；AGS 多了目的 Hub/Facility。但 Handover 冲突时，System 与 Platform-GS **完全一致**，AGS 是少数派。A-Scan、MAWB、Origin 高度一致。

System TikTok 独有 21,591 单：不在 AGS，全部在 Platform-GS。这与“TikTok 有多家清关行，System 含非 AGS 清关”的业务描述一致。

---

## 6. 重复单号字段对比

抽样规模：AGS-System 15、AGS-Platform 15、System-Platform 15、三方 15。以下统计基于**全量重叠**，样例仅作说明。

### 6.1 哪些字段基本一致

- **MAWB**：AGS-System、AGS-Platform、System-Platform 在双方都有值时 **冲突率 0**。
- **A-Scan Time**：三源双方都有值时 **冲突率 0**。Platform 还能补上少量 AGS/System 缺失的扫描时间（vs AGS +57，vs System +14）。
- **Origin Hub**：AGS 与 System / Platform **100% 一致**。
- **Dest Hub / Facility**：AGS 与 Platform-GS 完全一致；System 无这两列，不构成冲突。
- **System 与 Platform-GS 的 customer / MAWB / Handover / A-Scan**：几乎逐字段一致（Handover 冲突 0）。

### 6.2 哪些字段经常冲突

1. **客户名称写法（100% 字面冲突，实体相同）**  
   AGS `TIKTOK` vs System/Platform `TikTok Inc`。必须做别名，否则审计表会爆。

2. **Handover 时刻（同日不同时）**  
   - AGS vs System：2,266 / 11,612 = **19.5%**，全部同日历日。  
     样例：AGS `2026-08-24 17:16:00` vs System `2026-08-24 21:30:00`；或差 1 秒（`22:05:36` vs `22:05:35`）。  
   - AGS vs Platform：15,068 / 27,936 = **53.9%**。其中 TikTok 部分与上条同源；AliExpress 部分是 AGS 有秒 vs Platform 整分（如 AGS `14:20:41` vs Ali `14:20:00`，或 AGS `15:33:38` vs Ali `15:10:00`）。  
   - System vs Platform：**0%**。冲突时应以 System+Platform 为准，回写 AGS。

3. **Origin Hub（System vs Platform-GS，0.83%）**  
   277 单：System `IAH` vs Platform `DFW`（休斯顿线路 vs Dallas hub）。需要业务确认哪套 Hub 口径。

**未发现**同一单号对应完全不同客户实体、或不同 MAWB 的情况（在别名处理之前，客户只是命名差异）。

### 6.3 哪个来源在哪些字段上更完整

| 字段 | 更完整 / 更可信 |
|---|---|
| tracking number | 三源均完整 |
| customer | System / Platform-GS 用 `TikTok Inc`；AGS 需 join 且为 `TIKTOK`；Ali/FBT 无客户，只能用 AGS |
| MAWB | AGS 略全（TikTok 重叠中多 6 个）；FBT/Winit 无 MAWB |
| Pickup | **仅 System**（及待确认的 FBT Manifest / Winit manifested） |
| Handover | System = Platform-GS；AGS 有时偏差；Ali 精度较差 |
| A-Scan | 三源高度一致；Platform 可补少量空值 |
| Origin | 三源都较全；IAH vs DFW 需确认 |
| Dest Hub / Facility | **AGS 与 Platform 可补 System 的空缺** |
| 州 / 邮编 / 邮袋 | **仅 Platform-GS** |
| 24H 标记 / 时差 | 仅 AGS |

按 8 个标准字段非空计数（含 System 根本没有的 dest 列）：

- TikTok 重叠：AGS 平均 6.997 项 > System 4.997 → **AGS 字段更全，不是因为值更好，而是 schema 更宽**
- System ∩ Platform：Platform 6.998 > System 4.998 → 同样是 dest 列
- AGS ∩ Platform 且无 System（主要是 Ali）：AGS 6.993 > Platform 4.995 → **支持 AGS 主记录、Platform 补字段**

### 6.4 Platform 能否补充 AGS / System

**能，而且应该补充，不能因为重叠就丢弃。**

对 System ∩ Platform（33,203，全是 TikTok GS）：

- 可补 Dest Hub 33,203、Dest Facility 33,203
- 可补 A-Scan 14
- customer / MAWB / Handover **没有新值，但可作交叉验证**
- Pickup 双方都空（TikTok 无揽收）

对 AGS ∩ Platform（27,936 = GS 11,612 + Ali 16,324）：

- 可补 A-Scan 57
- Ali 侧可提供 `大包号`、`州`、`类型`
- GS 侧可提供 `州`、`邮编前三位`、`邮袋号`、`妥投时间`
- 客户：Ali 无客户，必须保留 AGS 客户（Cainiao 等）

Platform 独有 48,931 必须保留：FBT 24,994 + Winit 19,907 + GS 多出 2,143 + Ali 未进 AGS 1,887。

### 6.5 需要业务确认的冲突类型

- 客户别名是否视为同一客户（强烈建议是）
- Handover 同日不同时刻：以 System/Platform 还是 AGS
- Origin `IAH` vs `DFW`
- FBT Manifest=8/25、且部分 A-Scan 在 8/23，是否仍算 8/24 cohort
- Winit 701 条 8/25 是否纳入
- FBT 是否视为 TikTok

脱敏对照样例（三方重叠）：

| 单号（脱敏） | 字段 | AGS | System | Platform |
|---|---|---|---|---|
| SPXBWI****1265 | customer | TIKTOK | TikTok Inc | TikTok Inc |
| SPXBWI****1265 | handover | 22:05:35 | 22:05:35 | 22:05:35 |
| SPXRIC****2499 | handover | 22:05:36 | 22:05:35 | 22:05:35 |
| SPXMCO****6151 | handover | 21:08:00 | 21:33:55 | 21:33:55 |
| SPXPHX****4106 | dest_facility | （AGS 有对应值） | （空） | PHX |

---

## 7. 正式自动化方案（已按最终数据标准更新）

以下为已确认的开发方案。可用当前批次生成样例主明细与质量报告；每日自动化按同一规则运行。原始文件只读。


### 7.1 最终统一主明细字段顺序

主数据必须一单一行。对外主表明细 **V1.0 固定 17 列**，文件名为 `unified_master_YYYY-MM-DD.xlsx`（另存 CSV / Parquet）：

1. `Source`
2. `File名称`（全部分号连接）
3. `primary_source`（Platform → System → AGS）
4. `业务名称`
5. `TNO`
6. `MAWB`
7. `CUSTOMER`（标准化）
8. `customer_raw`（各来源原始 Customer，分号连接）
9. `Hub`
10. `start_time`
11. `start_time_type`（`Pickup Time` / `Handover Time` / `Manifest Time`）
12. `first_ascan_time`
13. `Ascan Due Date`（含时分秒）
14. `cohort_date`
15. `time_diff_hours`
16. `sla_hours`（TT FBT=12，其他=24）
17. `status`（仅 On Time / Late / No A-Scan，无 Pending）

`Ascan Due Date`：TT FBT 为 `start_time + 12 小时`；其余为 `+ 24 小时`。一律用当地时间，不用 UTC。无 A-Scan 一律 `No A-Scan`。

### 7.2 各来源字段映射

按文件结构、工作表和字段识别来源，不依赖文件名中的固定日期。

| Source | File名称 | 业务名称 | TNO | MAWB | Customer | Hub | 交接时间 | 交接时间类型 | Ascan时间 |
|---|---|---|---|---|---|---|---|---|---|
| AGS | `AGS_to_SPX_24H_Ascan_Report_YYYY-MM-DD` | AGS | Tracking Details `tracking_number` | Tracking Details `mawb` | 经 MAWB 匹配 MAWB Details `Customer` | MAWB Details `Hub` | Tracking Details `handoverTime_local` | Handover Time | Tracking Details `ascanTime` |
| System | `handover_pickup_YYYYMMDD` | 按 `customer_name` 填写 | `tracking_number` | `mawb` | `customer_name` | `origin_poe` | `pickup_poe_51005` 或 `handover_time_local` | Pickup Time 或 Handover Time | `first_ascan_poe`（实际为时间） |
| Platform | `平台客户需ASCAN单号明细 - TT-GS` | TT GS | `追踪号` | `主运单号` | TT GS | `Origin POE` | `交接当地时间` | Handover Time | `首次扫描当地时间` |
| Platform | `平台客户需ASCAN单号明细 - FBT` | TT FBT | `Tracking Number` | 空 | TT FBT | `Origin Hub` | `Manifest Time` | Manifest Time | `First Arrival Scan` |
| Platform | `平台客户需ASCAN单号明细 - Winit` | Winit | `Tracking Number` | 空 | Winit | `Origin Hub` | `date_manifested` | Manifest Time | `First Scan Date` |
| Platform | `平台客户需ASCAN单号明细 - AliExpress` | AliExpress | `跟踪号` | `提单号` | AliExpress | `Origin POE` | `清关交付-SPX收货时间` | Handover Time | `First Scan Arrival Scan` |

起算节点：

- AGS / TT-GS / AliExpress 清关：Handover Time
- System 本地揽收：Pickup Time
- FBT：Manifest Time（+12h）
- Winit：原始起算字段 `date_manifested`，主表 `交接时间类型` 显示为 Manifest Time（+24h）

### 7.3 AGS 文件内部处理顺序

1. 读取 `Tracking Details`
2. TNO 转文本、去空格、保留前导零
3. **先在 Tracking Details 内按 TNO 去重**，一单一行
4. 用去重后的 MAWB 匹配 `MAWB Details`
5. 补充 `Customer`、`Hub`
6. 若同一 MAWB 对应多个不同 Customer 或 Hub：不得任意选择；`mawb_mapping_conflict_flag = Yes`；写入冲突审计；主数据该字段暂时留空，并记录原因

空 TNO 进入 `invalid_no_tn`，不进主表。

### 7.4 跨来源去重与字段保留

以清洗后 TNO 为唯一主键，outer 合并，**禁止** append 后按单号删行。

字段级默认优先级：**Platform > System > AGS**。某来源该字段为空时，不得覆盖下一优先级的有效值。`Source` / `File名称` 仍保留全部来源。有 Platform 时 `primary_source = Platform`。

| 场景 | 规则 |
|---|---|
| 同一 TNO 只要 Platform 有数据 | 主表字段优先取 Platform；空字段再补 System、AGS |
| TikTok GS（Platform TT-GS明细） | `Customer` / `业务名称` = `TT GS`；交接时间、Ascan、Hub、MAWB 优先 Platform；不再显示 `TikTok Inc` / `TIKTOK` |
| Platform FBT / Winit / AliExpress | `Customer` / `业务名称` 分别为 `TT FBT` / `Winit` / `AliExpress` |
| AliExpress 与 AGS 重叠 | 主表 Customer = `AliExpress`；AGS 客户名（如 Cainiao）只留在审计原始字段 |
| AGS 原始 Customer 含 TikTok / TIKTOK / TikTok Inc | 主表 `Customer` / `customer_standardized` = `TT GS`；`customer_raw` 保留 AGS 原值 |
| AGS 原始 Customer = `HANGZHOU CAINIAO SUPPLY CHAIN MANAGEMENT CO.,LTD.` | 主表 `Customer` / `customer_standardized` = `AliExpress`；`customer_raw` 保留 Cainiao 全称 |
| Ascan 时间 | 与其他字段相同：Platform 有值用 Platform，否则 System，再否则 AGS |
| Hub 冲突 | Platform 优先；System/AGS 仅在 Platform Hub 为空时补充，并写审计 |
| 仅一个来源 | 保留该来源 |
| System 同时有 Pickup 与 Handover | `交接时间` 取时效起算节点（本批互斥；若未来同时存在：TikTok 用 Handover，否则 Pickup，并审计） |

`交接时间类型` 跟随最终采用的那条起算时间。Winit 原始字段名 `date_manifested` 不写入主表，统一显示 `Manifest Time`。

### 7.5 冲突审计与计算字段

冲突审计列：`TNO`、`field_name`、`ags_raw`、`system_raw`、`platform_raw`、`chosen_value`、`chosen_reason`、`needs_manual_review`、`primary_source`、`cohort_date`。

- `time_diff_hours` = Ascan时间 − 交接时间（小时）
- `within_24h` = 两者都有且时差 ≤ 24
- `monitor_flag`：缺起算、缺 A-Scan、或不在 24H 内 → Yes，否则 No

另输出 `off_day_records`（cohort_date 非主业务日）、`invalid_no_tn`。

---

## 8. 输入识别约定

按结构识别，不依赖 Platform 文件名日期：

| 类型 | 结构探测 |
|---|---|
| AGS | 存在 `Tracking Details` 与 `MAWB Details`，且含 `handoverTime_local` |
| System | 含 `pickup_poe_51005` 与 `handover_time_local` |
| Platform | 存在 `TT-GS明细` |

未上网宽表不计入主数据。`cohort_date` 一律来自 `交接时间` 本地日期。

---

## 9. 样例输出中请重点核对

1. 重叠 TikTok（Source 含 Platform）最终 Customer / 业务名称为 `TT GS`，不是 `TikTok Inc` / `TIKTOK`
2. AGS 独有 TikTok：`customer_raw` 仍为 `TIKTOK`，主表 `Customer` / `customer_standardized` = `TT GS`
3. AGS Cainiao：`customer_raw` 保留全称，主表 `Customer` = `AliExpress`
4. FBT/Winit 中 cohort_date=8/25 的记录仍在主表，同时进入 `off_day_records`
5. Hub 冲突采用 Platform，System/AGS 仅作空值补充，并出现在冲突审计
6. 主表 TNO 无重复
