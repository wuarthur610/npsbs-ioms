# NPSBS_IOMS_V1_0_Production_Contract_V1.0

**文件版本：V1.0**  
**文件日期：2026-09-15**  
**Owner：Arthur（亞瑟）**  
**Architect / Technical Gatekeeper：GPT / 夥伴**  
**Developer：Claude**  
**Independent QA：Gemini**  
**文件狀態：Draft for Claude + Gemini Review**

---

# 0. 文件目的

本文件是 NPSBS IOMS V1.0 的第一份正式 **Production Contract**。

目的不是描述「目前程式看起來怎麼寫」，而是定義：

> **什麼才算是 NPSBS IOMS V1.0 可以接受的正式系統。**

Claude 必須依本 Contract 實作。  
Gemini 必須依本 Contract 進行獨立 QA。  
GPT 負責架構與 Contract 的解釋、衝突裁決與 Acceptance Gate。  
Arthur 負責業務規則與最終 UAT。

---

# 1. 四方正式角色

| 角色 | 責任 |
|---|---|
| Arthur | Product Owner / Business Owner / Final UAT |
| GPT / 夥伴 | System Architect / Technical Gatekeeper |
| Claude | Developer / Implementer |
| Gemini | Independent QA / Code Reviewer |

核心原則：

> Claude 負責「怎麼實作」；Gemini 負責「是否可能有問題」；GPT 負責「是否符合架構契約」；Arthur 負責「業務是否正確」。

任何 AI 不得單方面改變 Arthur 已確認的業務規則。

---

# 2. Source of Truth 優先順序

當文件、AI 描述與實際程式互相矛盾時，依以下順序判斷：

1. 實際 `.xlsm` / `vbaProject.bin`
2. 實際 `.bas` / UserForm source
3. 實際 Workbook schema / sheet
4. 本 Production Contract
5. GPT Decision / Architecture MD
6. Claude Implementation Report
7. Gemini QA Report
8. 一般口頭描述

重要原則：

> **Documentation ≠ Source of Truth**

任何 AI 聲稱「已完成」都必須能指出實際 artifact。

---

# 3. Baseline / Candidate 定義

目前已驗證：

## FIXED14 — Verified Technical Baseline

已確認：

- 17 個 `.bas` Compile PASS
- BuildProductionUserForms PASS
- 8 個 UserForms 建立成功
- ApplyTraditionalChineseLabels PASS
- OpenMainMenu PASS
- frmTreatment 可開啟

FIXED14 不得被直接覆蓋。

---

## CANDIDATE-01

目前 Claude 所述的新候選版本：

```text
FIXED14
    ↓
PATCH-01
    ↓
CANDIDATE-01
```

CANDIDATE-01 必須在所有正式 Gate 通過前稱為：

> **Candidate**

不得稱：

> Production

---

# 4. 舊版 21-column Package

以下 artifact：

`NPSBS_IOMS_V1_MainSystem_Package.zip`

若其內容確為舊版：

- 21-column Treatment_Record
- `CreateTreatmentTransaction`

則正式標記：

> **OBSOLETE / DO NOT USE**

不得作為目前開發基準。

但不得刪除。應保留於 Archive，以維持完整版本歷史。

---

# 5. Business Rules

## 5.1 Calendar

Google Calendar 是：

> **Actual Treatment Schedule / Actual Time Source**

Excel 每日預約表是：

> **HQ / Reporting / Upload Format**

因此：

> Appointment ≠ Treatment

Walk-in、電話、LINE、臨時變更均可能造成兩者不同。

---

# 6. Calendar Cancellation / Reschedule

取消：

- 原 Calendar event 保留。
- 原事件改 graphite black。
- 不覆蓋原歷史。

改期：

- 建立新日期 / 時間的 event。
- 原 event 保留。
- 不將歷史直接改寫。

---

# 7. Treatment Rules

目前服務：

> 淋巴紓壓調理

標準：

- 40 分鐘 / session
- 一般 2 sessions
- 特殊情況 1 或 3 sessions

Treatment_Record 必須保存：

- 實際治療日期
- 實際開始時間
- 實際結束時間
- 調理師
- session 數
- booking / treatment 狀態
- Card 使用狀態
- shortage 狀態
- notes
- audit metadata

---

# 8. Card / Package Master

目前有效 package：

| Card Type | Purchase Amount | C1 Qty | Sessions | Product Discount |
|---|---:|---:|---:|---:|
| 彩綿卡 | 24,000 | 6 | 16 | 8 折 |
| 華夏尊盈 | 60,000 | 5 | 35 | 7 折 |
| 闔家萬吉 | 100,000 | 10 | 60 | 6 折 |
| 映遊 | 16,000 | 4 | 10 | 9 折 |

不得從其他舊版資料自行覆蓋以上正式定義。

---

# 9. CardID Rule

正式規則：

```text
彩綿 → C + YYYYMMDD + ###
華夏 → H + YYYYMMDD + ###
闔家 → W + YYYYMMDD + ###
映遊 → TBD
```

例如：

```text
C20260915001
H20260915001
W20260915001
```

### 重要

映遊 prefix 尚未正式確認。

不得猜測。

CardID 產生邏輯必須集中於 `modCard` / Card business layer。

不得讓不同 UserForm 各自建立 prefix。

---

# 10. Card Usage Rules

正常情況：

- 客戶通常一次使用一張卡。
- 舊卡仍有餘額時，不建立第二張卡作為正常扣款來源。
- 不可把過去 Treatment retrospectively 扣到未來購買的卡。
- RemainingSessions 不得小於 0。

Card deduction 必須留下：

- TreatmentRecordID
- CardID
- UsageDate
- BeforeSessions
- UsedSessions
- AfterSessions
- UsageType / Notes
- CreatedAt

---

# 11. Shortage Rules

Treatment 所需 sessions 超過卡片剩餘堂數時，可以：

### Option A
收取減少堂數費用。

### Option B
直接減免。

### Option C
Pending Settlement：

暫不扣足，待下一次購卡再補扣。

---

# 12. Pending Settlement

Pending Settlement 不得：

- 修改原 Treatment 成為另一筆治療
- 直接把舊 Treatment 改寫成新卡扣款

應：

1. 保留原 Treatment_Record。
2. 保留原 shortage 狀態。
3. 下一次結算建立新的 Card_Usage_Detail。
4. `UsageType = PendingSettlement`。
5. Audit 保留完整關聯。

---

# 13. Card Balance Guard

所有 Card_Master 更新前必須驗證：

```text
BeforeSessions >= UsedSessions
```

且：

```text
AfterSessions >= 0
```

任何流程不得產生：

```text
RemainingSessions < 0
```

`On Error Resume Next` 不得被視為 business validation。

Business validation 必須明確存在。

---

# 14. Customer Rules

CustomerID：

- 系統自動產生。
- Manager 可人工調整。
- 調整必須留下 Audit。

FirstVisitDate：

- 預設依第一筆 actual treatment 自動產生。
- Manager 可以人工修正。
- 原始值與修改歷史必須保留。

CustomerName 是一般操作的主要搜尋入口。

---

# 15. Privacy Boundary

`客戶資料表.pdf` 含較敏感的個人 / 健康資訊。

因此：

- 不應混入一般每日治療操作流程。
- 不應讓一般 daily treatment UI 任意暴露敏感欄位。
- M3-14 保留，但不再擴大開發。

---

# 16. Treatment_Record Schema Contract

目前正式 Contract 必須以 **22-column candidate schema** 為基準。

在 Claude 提交 CANDIDATE-01 的 source artifact 後，必須填入實際：

- Column
- Field Name
- Data Type
- Required
- Source
- Validation
- Description

### 特別要求

不得同時存在：

```text
CreateTreatmentTransaction
```

與

```text
CreateTreatmentRecord
```

兩套不同 Production transaction API 而未明確標示版本。

正式 Production API 應鎖定單一 transaction entry point：

```text
CreateTreatmentRecord(...)
```

若舊 function 仍存在，必須：

- 明確標記 legacy
- 不由 Production UI 呼叫
- 或移除 / deprecated

---

# 17. Card_Master Schema Contract

至少必須支持：

- CardID
- CustomerID
- CardType
- PurchaseDate
- PurchaseAmount
- OriginalSessions
- UsedSessions
- RemainingSessions
- PaymentMethod
- Active
- CreatedAt
- CreatedBy
- ModifiedAt
- ModifiedBy

實際欄位順序與額外欄位需由 Candidate artifact + schema manifest 最終鎖定。

---

# 18. Card_Usage_Detail Schema Contract

至少必須支持：

- TreatmentRecordID
- CardID
- UsageDate
- BeforeSessions
- UsedSessions
- AfterSessions
- UsageType
- Notes
- CreatedAt
- CreatedBy

正式 production source 必須能回答：

> 「這一次扣堂是因為哪一次 Treatment？」

以及：

> 「扣堂前後卡片還剩多少？」

---

# 19. Card_Product_Detail

規則：

- CardID 為主要關聯。
- Detail 可以由 Manager 建立 / 修正。
- 修改不得破壞舊版本。
- 舊版本需保留。
- CardID 可以協助判斷產品出貨內容。

---

# 20. Inventory

Inventory 採簡化 reconciliation：

1. 盤點。
2. 發現差異立即重盤。
3. 確認差異原因。
4. Difference Reason 必須選擇。
5. 不允許 unexplained discrepancy。
6. `Loss` 可以為合法原因。
7. Reconciliation 完成後才結案。

---

# 21. UserForm Architecture

Production UI 必須遵守：

```text
UserForm / Presentation
        ↓
Application / Business Logic
        ↓
Data / Transaction Layer
        ↓
Worksheet
```

不建議將大量 business transaction logic 直接塞入 `cmdSave_Click`。

例如 Treatment：

```text
frmTreatment.cmdSave_Click
        ↓
SaveTreatmentForm(...)
        ↓
Validate
        ↓
CreateTreatmentRecord(...)
        ↓
Card / Usage / Audit
```

---

# 22. UserForm Control Contract

Control name 必須是 contract 的一部分。

例如：

```text
frmTreatment
    txtCustomerName
    txtCustomerID
    txtTreatmentDate
    cboTherapist
    txtStartTime
    txtEndTime
    txtTreatmentType
    txtBookingType
    txtSessions
    ...
```

Claude 每次修改 Form 必須提供：

> Control Manifest

Gemini 必須檢查：

```text
Control exists
    +
Handler references same name
    +
Data type matches
    +
Save logic maps correctly
```

---

# 23. Control Mismatch QA Rule

如果 Gemini 發現：

```text
Build:
cboTreatmentType
```

但 handler 使用：

```text
txtTreatmentType
```

必須提供：

- artifact
- module
- procedure
- exact line / source snippet
- reproduction status

若實際 source 證明兩者一致：

> Finding 必須標記 `REJECTED / NOT REPRODUCIBLE`

不得只因舊版本曾經出現過而維持 Critical。

---

# 24. Transaction Integrity

Treatment + Card deduction 是核心交易。

理想流程：

```text
START
  ↓
Validate Input
  ↓
Validate Customer
  ↓
Validate Therapist
  ↓
Validate Card
  ↓
Validate Sessions
  ↓
Snapshot Card
  ↓
Write Treatment_Record
  ↓
Write Card_Usage_Detail
  ↓
Update Card_Master
  ↓
Audit
  ↓
COMMIT
```

---

# 25. Rollback Requirement

如果任一步失敗：

```text
ERROR
  ↓
Rollback Treatment_Record
  ↓
Rollback Card_Usage_Detail
  ↓
Restore Card_Master
  ↓
Write Error / Audit
```

Excel/VBA 沒有傳統 database transaction，因此 Production 至少必須具備：

> **Pseudo-Transaction / Explicit Rollback**

不能只依賴：

```text
On Error GoTo EH
```

就宣稱具備 transaction safety。

---

# 26. Atomicity Acceptance Test

Gemini / GPT 必須設計至少以下 failure tests：

### Test A
Treatment_Record write failure。

預期：

- Card 不應永久少堂。
- 不應留下不完整 Treatment。

### Test B
Card_Master update failure。

預期：

- Treatment / Usage 不應留下不一致狀態。

### Test C
Card_Usage_Detail failure。

預期：

- Card_Master 與 Treatment 必須 rollback 或保持一致。

---

# 27. Audit Contract

重大變更必須留下：

- Who
- When
- What
- Old Value
- New Value
- Reason（適用時）

尤其：

- CustomerID
- FirstVisitDate
- Card
- Treatment
- Card Product Detail
- Inventory reconciliation
- Manual correction

---

# 28. Historical Migration

歷史範圍：

```text
2025/12 ～ 2026/08
```

已知：

- Calendar Events = 731
- Excel Candidates = 416
- High Confidence Matches = 248
- Review = 156
- Unmatched = 12

Production import 必須經 UAT Gate。

不得因 migration 自動覆蓋目前 Master Data。

---

# 29. September 2026 UAT

2026/09/01–09/07：

- Calendar = 23 events
- Daily Excel = 15 treatment records
- Sessions = 30

已知 discrepancy：

`卜詠妍 (Alex)` 9/2 Calendar 有治療事件，但 Daily Excel 沒有。

因此：

> 必須保持 Review / Staging，未經確認不得自動進 Production Treatment_Record。

Revenue cells = 0：

> 不得自行從 appointment 推導 revenue。

---

# 30. QA Evidence Standard

Gemini 所有重要 finding 必須包含：

```text
Severity
Artifact
Version
Module
Procedure
Line / Source
Observed Behavior
Expected Behavior
Reproduction
Recommendation
```

Severity：

```text
CRITICAL
HIGH
MEDIUM
LOW
CONFLICT
UNVERIFIED
```

### 特別規則

如果沒有足夠 source evidence：

> 必須使用 `[UNVERIFIED]`

不能直接升級為 `[CRITICAL]`。

---

# 31. QA 不可把推測當成事實

以下類型必須標示：

```text
[UNVERIFIED]
```

例如：

- 推測 VBE cache 問題
- 推測某個錯誤一定由 On Error Resume Next 造成
- 推測某個 control 在另一版本不存在

必須區分：

> Observed Fact

與：

> Hypothesis

---

# 32. Version Evidence

每次 Candidate 必須提供：

```text
Candidate Name
Date
Source Artifact
File Size
SHA-256（建議）
Modified Modules
Unmodified Modules
Schema Version
Build Result
Compile Result
```

若可能，建立：

`MANIFEST_SHA256.txt`

---

# 33. Definition of Done

「完成」不等於 Claude 寫完。

正式狀態：

| Status | Definition |
|---|---|
| Designed | 設計完成 |
| Implemented | Claude 已實作 |
| Static QA Passed | 靜態檢查通過 |
| Compile Passed | Excel Compile 通過 |
| Runtime Passed | 實際執行通過 |
| Gemini QA Passed | Gemini QA 通過 |
| GPT Gate Passed | GPT 架構 Gate 通過 |
| Arthur UAT Passed | Arthur 業務驗收通過 |
| Production Candidate | 全部前置 Gate 通過 |
| Production | Arthur 正式核准 |

---

# 34. Patch Workflow

正式流程：

```text
Production Contract
        ↓
GPT Patch Specification
        ↓
Claude Implementation
        ↓
Claude Implementation MD
        ↓
Gemini Independent QA
        ↓
Gemini QA MD
        ↓
GPT Architecture Decision
        ↓
Arthur UAT
        ↓
New Baseline
```

禁止：

```text
Gemini finding
    ↓
Claude immediate wholesale rewrite
```

---

# 35. Decision Codes

GPT Decision MD 必須使用：

```text
ACCEPT
REJECT
PATCH REQUIRED
DEFER
BLOCK
```

Arthur Business Decision：

```text
APPROVE
REJECT
CHANGE REQUIREMENT
```

---

# 36. AI Handoff MD Standard

所有重要 AI 回覆必須同時提供 MD。

MD 至少包含：

```markdown
# Executive Summary
# Version / Baseline
# Files Examined
# Confirmed Facts
# Findings
# Evidence
# Decisions
# Open Issues
# Recommended Action
# Next AI Action
```

Chat：

> 給 Arthur 看。

MD：

> 給下一個 AI 工作。

---

# 37. No Manual Copy/Paste Principle

Arthur 不應成為 AI 之間的人工資料搬運者。

優先：

```text
AI
 ↓
MD
 ↓
下一 AI
```

而不是：

```text
AI
 ↓
Arthur Copy
 ↓
Paste
 ↓
另一 AI
```

如需人工決策，才由 Arthur 介入。

---

# 38. Current Priority

本 Contract 通過後，下一優先順序：

### P0
鎖定 CANDIDATE-01 實際 artifact。

### P1
完成 Treatment_Record 22-column schema manifest。

### P2
完成 exact function signatures。

### P3
完成 UserForm Control Manifest。

### P4
完成 Transaction / Rollback design。

### P5
Claude 實作 PATCH-02。

### P6
Gemini QA-02。

### P7
Arthur UAT。

---

# 39. Current Known Issues

## Issue 001 — 21 vs 22 Column

Status：

`RESOLVED AS VERSION CONFLICT`

舊 21-column package = obsolete。

CANDIDATE-01 = 22-column candidate，仍需以 actual artifact 驗證。

---

## Issue 002 — Treatment Control mismatch

Gemini 曾報告：

`cboTreatmentType` vs `txtTreatmentType`

Claude 最新 evidence 顯示：

`txtTreatmentType` 一致。

Current status：

`REJECTED / NOT REPRODUCIBLE`

除非新的 actual artifact 再次證明。

---

## Issue 003 — Atomicity

Status：

`OPEN / HIGH`

必須進入 PATCH-02。

---

## Issue 004 — CardID / 映遊 Prefix

Status：

`OPEN / MEDIUM`

映遊 prefix 尚未確認。

---

## Issue 005 — VBE cache / DoEvents

Status：

`UNVERIFIED`

不得當成正式 root cause。

---

# 40. Acceptance Gate

CANDIDATE-01 不能進 Production Candidate，除非：

- [ ] Schema locked
- [ ] Function signatures locked
- [ ] Control manifest verified
- [ ] Compile PASS
- [ ] Runtime PASS
- [ ] Card balance guard PASS
- [ ] Atomicity test PASS
- [ ] Rollback test PASS
- [ ] Gemini QA PASS
- [ ] GPT Architecture Gate PASS
- [ ] Arthur UAT PASS

---

# 41. Arthur 的最終權限

任何 AI 都不能代替 Arthur：

- 改變商業規則
- 調整 package 價格
- 調整 sessions
- 確認未知 CardID prefix
- 決定 Production release

Arthur 是最終 Business Owner。

---

# 42. Final Principle

NPSBS IOMS V1.0 不追求：

> 「看起來可以跑。」

而追求：

> **Business Correctness + Data Integrity + Auditability + Maintainability + QA Reproducibility + UAT Acceptance**

Production Candidate 必須通過：

```text
Business Rule
      ↓
Contract
      ↓
Source Evidence
      ↓
Compile
      ↓
Runtime
      ↓
Data Integrity
      ↓
Transaction Integrity
      ↓
Gemini QA
      ↓
GPT Gate
      ↓
Arthur UAT
      ↓
Production
```

---

# 43. Review Request — Claude

請 Claude：

1. 以本 Contract 為正式 review baseline。
2. 提供 CANDIDATE-01 實際 artifact。
3. 提供 22-column Treatment_Record schema manifest。
4. 提供所有 Production transaction function signatures。
5. 提供 UserForm Control Manifest。
6. 提供 modified/unmodified module manifest。
7. 不進行 Contract 之外的 wholesale rewrite。
8. 對任何無法實作的條款提出 `[CONFLICT]`，不要自行猜測。

---

# 44. Review Request — Gemini

請 Gemini：

1. 以本 Contract 為 QA baseline。
2. 使用實際 CANDIDATE-01 artifact。
3. 每一項 finding 提供 evidence。
4. 區分 Fact / Hypothesis。
5. 無法證明時標 `[UNVERIFIED]`。
6. 優先驗證：
   - schema
   - signatures
   - controls
   - transaction
   - card balance
   - rollback
   - audit
7. 不進行 wholesale rewrite。
8. 輸出 Gemini QA MD。

---

# 45. Review Request — GPT

GPT / 夥伴負責：

1. 整合 Claude implementation。
2. 整合 Gemini QA。
3. 判定 Contract compliance。
4. 產生 Architecture Decision。
5. 管理 baseline。
6. 決定是否進入下一 Patch。
7. 最終將技術結果翻譯成 Arthur 看得懂的中文。

---

# 46. 文件狀態

**Status：DRAFT FOR REVIEW**

本文件送交：

- Claude — Developer Review
- Gemini — Independent QA Review

兩者回覆後，由 GPT 產生：

> `NPSBS_IOMS_V1_0_Production_Contract_V1.1.md`

只有 V1.1 通過後，才作為正式 Production Contract。

---

**END OF CONTRACT**
