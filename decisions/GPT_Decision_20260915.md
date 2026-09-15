# GPT Architecture Decision — NPSBS IOMS V1.0

**Decision ID:** GPT-DEC-20260915-01  
**Decision Date:** 2026-09-15  
**Project:** NPSBS IOMS V1.0  
**Decision Authority:** GPT — System Architect / Architecture Gate  
**Product Owner:** Arthur  
**Developer:** Claude  
**Independent QA:** Gemini  

---

## 1. Decision Status

# CONDITIONAL APPROVAL

**Architecture direction: APPROVED**

**Production release: NOT APPROVED**

本次裁決確認 NPSBS IOMS V1.0 可以繼續進入下一階段開發，但目前 Candidate 尚未達到 Production Release Gate。

Production Release 必須等待以下事項完成：

1. Production Contract V1.1 鎖定。
2. Claude 完成 Contract V1.1 所要求的 P0 修正。
3. Gemini 完成獨立 QA / line-by-line cross-check。
4. Excel `Debug → Compile VBAProject` 通過。
5. UserForm 建置與中文標籤套用通過。
6. Treatment / Shortage / Pending Settlement / Rollback UAT 通過。
7. GPT 完成最終 Architecture Gate。
8. Arthur 完成最終 Product Owner UAT。

---

# 2. Decision Scope

本次裁決針對：

- FIXED14 technical baseline
- CANDIDATE-01
- Claude implementation direction
- Gemini QA findings
- Treatment_Record schema
- Treatment transaction architecture
- CardID generation
- Card selection
- Card_Usage_Detail
- Shortage / Pending Settlement
- UserForm control contract
- Production Contract V1.1

進行架構與工程層級裁決。

本文件為 **Architecture Decision Record**，不是 Developer Implementation Guide。

---

# 3. Source-of-Truth Hierarchy

若不同文件、Package、AI handoff 或舊版本描述互相衝突，採以下優先順序：

1. Approved Production Contract
2. GPT Architecture Decision
3. 實際 GitHub source code
4. 實際 `.xlsm` / `.xlsx` binary evidence
5. Claude implementation report
6. Gemini QA report
7. 舊版 Markdown / handoff / package description

**實際 source artifact 優先於文件描述。**

任何 AI 交接文件若與實際 source code 不一致，不得直接視為目前實作狀態。

---

# 4. Baseline Decision

## 4.1 FIXED14

**Status: VERIFIED BASELINE**

FIXED14 為目前已經由 Arthur 實際完成以下驗證的 technical baseline：

- 17 VBA modules imported
- `Debug → Compile VBAProject` = PASS
- `BuildProductionUserForms` = PASS
- 8 production UserForms successfully created
- `ApplyTraditionalChineseLabels` = PASS
- `OpenMainMenu` = PASS
- `frmTreatment` successfully opened

因此 FIXED14 可作為後續 Candidate 的 technical comparison baseline。

---

## 4.2 CANDIDATE-01

**Status: CANDIDATE ONLY**

CANDIDATE-01 可以作為下一階段開發基礎，但目前不得標記為 Production。

原因：

- transaction atomicity 尚未完成；
- CardID format 尚需修正；
- Card_Usage_Detail balance traceability 尚需對齊；
- oldest-card selection 尚需修正；
- pending settlement atomicity 尚需完成；
- 尚待 Gemini QA；
- 尚待 Arthur Excel compile / UAT。

---

## 4.3 Legacy 21-Column MainSystem

**Status: OBSOLETE / ARCHIVE**

Claude 舊版 MainSystem package 所使用的 21-column Treatment_Record implementation，不得再作為目前 Production source。

該版本可保留作為歷史 / archive evidence，但不得與目前 22-column Production target 混用。

---

# 5. Treatment_Record Schema Decision

## Decision

**22-column schema = APPROVED PRODUCTION TARGET**

正式欄位如下：

| # | Field |
|---:|---|
| 1 | TreatmentRecordID |
| 2 | TreatmentDate |
| 3 | TreatmentStart_Taipei |
| 4 | TreatmentEnd_Taipei |
| 5 | CustomerName |
| 6 | CustomerID |
| 7 | Therapist |
| 8 | TherapistID |
| 9 | Sessions |
| 10 | TreatmentType |
| 11 | BookingType |
| 12 | CardID |
| 13 | CardUsageStatus |
| 14 | ShortageResolution |
| 15 | ShortageSessions |
| 16 | TreatmentRevenue |
| 17 | PaymentMethod |
| 18 | CalendarUID |
| 19 | Notes |
| 20 | CreatedAt |
| 21 | CreatedBy |
| 22 | Status |

---

# 6. Gemini CRITICAL-01 Decision

## Finding

Gemini identified a potential conflict between:

- 22-column Treatment_Record
- `CreateTreatmentRecord(...)` parameter count

## GPT Decision

**REJECTED AS A CONTRACT VIOLATION**

原因：

Database columns 不要求與 VBA function parameters 一對一對應。

下列欄位可以由 transaction layer 自動產生：

- CustomerID
- CardUsageStatus
- ShortageSessions
- CreatedAt
- CreatedBy
- Status

因此：

> `22 database columns ≠ 22 VBA parameters`

不是 architecture defect。

---

# 7. Gemini CRITICAL-02 Decision

## Finding

Gemini reported a possible mismatch:

- `cboTreatmentType`
- `cboBookingType`

vs.

- `txtTreatmentType`
- `txtBookingType`

## GPT Decision

**REJECTED / NOT REPRODUCIBLE AGAINST CURRENT CANDIDATE**

目前 Candidate source 使用：

- `txtTreatmentType`
- `txtBookingType`

而 save path 亦使用相同 control names。

因此目前沒有足夠 evidence 證明存在 Gemini 所描述的 control mismatch。

## Governance Rule

未來所有 UserForm QA 必須以：

> actual `modFormBuilder.bas` + actual form event code

為準，而不是舊版報告或推測。

---

# 8. P0 Decision — Treatment Transaction Atomicity

## Finding

目前 Candidate 的 transaction flow 存在：

```text
Validate
    ↓
Find Card
    ↓
Deduct Sessions
    ↓
Write Treatment_Record
    ↓
Create Card_Usage_Detail
    ↓
Handle Shortage
