# NPSBS_IOMS_V1_0_Claude_Production_Contract_Review_V1.0

**文件版本**：V1.0-Claude-Review-20260915
**文件日期**：2026-09-15
**作者 AI**：Claude（Developer）
**審閱對象**：`NPSBS_IOMS_V1_0_Production_Contract_V1.0.md`（Status: DRAFT FOR REVIEW）
**文件用途**：依 Contract §43 提出 Developer Review。**本輪未修改任何程式碼**，
依 Arthur 指示僅做 Contract Review。
**Review baseline**：CANDIDATE-01（17 個 `.bas`）+ `NPSBS_IOMS_V1_0.xlsm` 實際工作表

---

# Executive Summary

**整體立場：同意 Contract 的架構方向，建議通過為 V1.1，但有 4 個 `[CONFLICT]` 必須先解決，
其中 2 個是 Contract 條文與實際 artifact 直接矛盾、無法同時成立。**

同時，在這次 Review 過程中發現 **1 個 `[CRITICAL]` 等級的既有缺陷**，
它現在就存在於 Arthur 已驗證的 FIXED14 中，而且正好是 Contract §26 Test A 想防的那一類問題。
這個缺陷**會造成重複扣堂**，建議列為 PATCH-02 的最高優先。

我也接受 Contract 對我先前一項建議的糾正（§4 保留 obsolete package 而非刪除），
並主動申報我自己 PATCH-01 的一個實作缺口（CONFLICT-01）。

---

# Version / Baseline

| 項目 | 值 |
|---|---|
| Review 對象 Contract | V1.0 / 2026-09-15 / DRAFT |
| Claude 交付 Candidate | CANDIDATE-01 / 2026-09-11 |
| Candidate 基底 | FIXED13 `.bas` ×17（= FIXED14 內容） |
| 修改模組 | `modCard`、`modFormBuilder`、`modText`（3 個） |
| 未修改模組 | 其餘 14 個，與 FIXED13 逐字節相同 |
| Compile 狀態 | **UNVERIFIED** — Arthur 尚未回報 Compile 結果 |

**重要**：CANDIDATE-01 目前仍停留在 Contract §33 的 `Implemented` + `Static QA Passed`，
**尚未達到 `Compile Passed`**。本 Review 中所有關於 CANDIDATE-01 的陳述都以此為前提。

---

# Files Examined

| 檔案 | 檢查方式 |
|---|---|
| `NPSBS_IOMS_V1_0_Production_Contract_V1_0.md` | 全文逐節閱讀 |
| `NPSBS_IOMS_V1_0.xlsm` 全部工作表第 1 列 | `openpyxl` 讀取實際表頭與資料列數 |
| CANDIDATE-01 `modTreatment.bas` | 逐行閱讀 `CreateTreatmentRecord` 執行順序 |
| CANDIDATE-01 `modCard.bas` | 逐行閱讀 `HandleShortage`、`DeductSessions` |
| CANDIDATE-01 `modCardUsage.bas` | 逐行閱讀 `CreateUsageDetail` 實際寫入欄位 |
| CANDIDATE-01 `modValidation.bas` | 逐行閱讀 `ValidateShortageResolution` |
| CANDIDATE-01 `modFormBuilder.bas` | 控制項名稱清單 vs Contract §22 |

---

# Confirmed Facts（實際 artifact 證據）

## F-1 實際工作表 schema（`NPSBS_IOMS_V1_0.xlsm`，Contract §2 第 3 順位 source of truth）

| 工作表 | 欄數 | 資料列數 | 實際欄位 |
|---|---:|---:|---|
| `Treatment_Record` | 22 | 0 | 與 Contract §16 完全一致 |
| `Card_Master` | 10 | **24** | CardID, CustomerID, PurchaseDate, PurchaseAmount, OriginalSessions, UsedSessions, RemainingSessions, PaymentMethod, CardStatus, Notes |
| `Card_Usage_Detail` | 12 | 0 | UsageDetailID, TreatmentRecordID, CardID, CustomerID, UsageDate, **SessionsUsed**, UsageType, ShortageSessions, Notes, CreatedAt, CreatedBy, Version |
| `Customer_Master` | 5 | 89 | CustomerID, CustomerName, FirstVisitDate, Notes, Status |
| `Pending_Shortage` | 11 | 0 | 見 Contract §12 相關欄位，齊備 |
| `Card_Product_Detail` | 13 | 0 | 含 VersionNo / EffectiveFrom / EffectiveTo，**已支援 §19 版本保留** |
| `Audit_Log` | 10 | 0 | AuditID, AuditTime, UserName, ActionType, TableName, RecordID, FieldName, OldValue, NewValue, Reason |

`Treatment_Record` 22 欄與 Contract §16 一致 → **Issue 001 可正式關閉。**

## F-2 `CreateTreatmentRecord` 的實際執行順序（CANDIDATE-01 `modTreatment.bas`）

```
1. EnsureProductionSheets
2. GetCustomerByName → 無則 raise
3. ValidateTreatment / ValidatePaymentMethod
4. FindUsableCard 或 ValidateCardForCustomer
5. DeductSessions(actualCard, sessions, shortage)   ← 【此處已寫入 Card_Master】
6. 寫入 Treatment_Record（22 欄）
7. CreateUsageDetail
8. HandleShortage → 內部第一行才呼叫 ValidateShortageResolution
9. WriteAuditLog
EH: Err.Raise（僅重拋，無任何 rollback）
```

---

# Findings

## `[CRITICAL]` FIND-01 — 不足堂數未選處理方式時，產生半套交易且重試會重複扣堂

| 欄位 | 內容 |
|---|---|
| **Severity** | CRITICAL |
| **Artifact** | CANDIDATE-01（同時存在於 FIXED14，**非 Claude PATCH-01 引入**） |
| **Module / Procedure** | `modTreatment.CreateTreatmentRecord` + `modCard.HandleShortage` + `modValidation.ValidateShortageResolution` |
| **Evidence** | `modTreatment.bas`：第 5 步 `DeductSessions` → 第 6 步寫 `Treatment_Record` → 第 8 步才 `HandleShortage`。<br>`modCard.bas` `HandleShortage` 第 1 行：`ValidateShortageResolution shortage,resolution`。<br>`modValidation.bas` 第 22–25 行：`If shortage <= 0 Then Exit Sub` / `Case Else: Err.Raise ... MsgText("SHORTAGE_RESOLUTION")`。 |
| **Observed Behavior** | 當 `shortage > 0` 且店長未選處理方式時：(1) `Card_Master` 已被實際扣堂；(2) `Treatment_Record` 已寫入且 `ShortageResolution` 為空白；(3) 此時才拋出「不足堂數必須選擇處理方式」；(4) `EH` 只 `Err.Raise`，**不回滾任何已寫入的資料**。店長看到錯誤訊息，合理認為「沒有存檔」，於是補選處理方式再按一次儲存 → **`DeductSessions` 第二次執行，同一次治療被扣兩次堂數，並產生第二筆 `Treatment_Record`。** |
| **Expected Behavior** | 依 Contract §24，所有 validation 必須在任何寫入之前完成；依 §26 Test A，卡片不應永久少堂。 |
| **Reproduction** | 尚未實機執行（Claude 無 Excel 環境），但為**純靜態可推導的執行路徑**，非假設。判定為 Observed Fact（source-level），Runtime 驗證請 Gemini / Arthur 補做。 |
| **Recommendation** | PATCH-02 最高優先。最小修正：把 `ValidateShortageResolution` 與 `TW_Charge` 的 `shortageAmount > 0` 檢查，前移到 `DeductSessions` 之前（此修正不需重構架構，屬 §24「Validate Input」階段）。 |

> 同一路徑另有一條較輕的變體：選 `TW_Charge()` 但 `shortageAmount <= 0` 時，
> `HandleShortage` 第 3 行才 raise，同樣已完成扣堂與寫入。

---

## `[CONFLICT]` 需 GPT 裁決的 4 項

### CONFLICT-01 — §17 Card_Master 欄位要求 vs 實際 10 欄（含 Claude 自身缺口申報）

Contract §17 要求至少支援 14 個欄位。實際 `Card_Master` 只有 10 欄，缺少：

| Contract 要求 | 實際狀態 |
|---|---|
| `CardType` | **不存在** |
| `Active` | 實際為 `CardStatus`（語意可對應，命名不同） |
| `CreatedAt` / `CreatedBy` | **不存在** |
| `ModifiedAt` / `ModifiedBy` | **不存在** |
| （實際有但 §17 未列）`Notes` | 需確認是否保留 |

**Claude 主動申報自身缺口**：PATCH-01 的 `CreateCard` 已新增 `cardType` 參數，
但因實體欄位不存在，**該參數只用於推導 CardID 前綴，並未寫入任何欄位**。
後果是：卡種資訊只隱含在 CardID 第一個字元裡，一旦手動輸入 CardID 就完全遺失，
**無法稽核、無法用於 §19 判斷應出貨產品**。這是 PATCH-01 的實作缺口，我應該在交付時就指出，
當時沒有，在此補正。

**衝突點**：`Card_Master` 有 **24 筆歷史資料**（`CardStatus = Observed`，備註「待紙本套卡購買紀錄核對」）。
新增欄位後這些歷史列的新欄位皆為空白，而 §28 規定「不得因 migration 自動覆蓋目前 Master Data」，
§8 規定「不得從其他舊版資料自行覆蓋正式定義」。

**Claude 建議**：新增 `CardType`、`CreatedAt`、`CreatedBy`、`ModifiedAt`、`ModifiedBy` 五欄（附加在第 11–15 欄，
不改動既有 1–10 欄位置，確保既有程式的欄位索引全部不受影響）；歷史 24 列的 `CardType` 留白，
不回填、不猜測，等紙本核對。`Active` 建議維持實際命名 `CardStatus`，由 Contract 改名以對齊 artifact
（§2 規定 artifact 優先於 Contract）。

**Decision Needed（GPT）**：確認欄位新增方式與 `Active` / `CardStatus` 的命名裁決。

---

### CONFLICT-02 — §18 要求 BeforeSessions / AfterSessions，但實際 schema 沒有，且結構上無法回答 §18 自己提的問題

Contract §18 要求 `Card_Usage_Detail` 至少支援 `BeforeSessions`、`UsedSessions`、`AfterSessions`，
並明確要求這張表必須能回答：

> 「扣堂前後卡片還剩多少？」

**實際 `Card_Usage_Detail` 只有 `SessionsUsed`，沒有 `BeforeSessions`，也沒有 `AfterSessions`。**
`modCardUsage.CreateUsageDetail` 的 `EnsureSheet` 陣列逐字列出 12 個欄位，其中不含這兩者。

因此**目前 schema 在結構上無法滿足 §18 的驗收要求**。

**需要向 GPT 指出的一條追溯線**：GPT 在
`NPSBS_IOMS_V1_0_GPT_Latest_Progress_Handoff_20260911.md` §20 曾寫下
「Claude 使用 TreatmentRecordID / CardID / UsageDate / **BeforeSessions** / UsedSessions /
**AfterSessions** / Notes / CreatedAt……GPT 認同這個歷史明細設計」。

該設計出自 **已作廢的 21-column MVP 草稿**，不是 FIXED14。
換言之，§18 很可能是把作廢草稿的設計誤寫進了 Contract。這與 Issue 001 同源，
是同一次版本混淆的第二個殘留。

**Claude 建議（二擇一，由 GPT 裁決）**：
- **方案 A**：`Card_Usage_Detail` 新增 `BeforeSessions`、`AfterSessions` 兩欄（附加於末尾，不影響既有索引）。
  優點：符合 §18 原意，單表即可完整回答；此表目前 0 筆資料，新增零風險。
- **方案 B**：維持現有 schema，改由 `Audit_Log` 回答（`DeductSessions` 目前已寫入
  `FieldName="RemainingSessions"`、`OldValue`、`NewValue`，資訊其實存在）。
  修改 §18 條文為「由 Card_Usage_Detail + Audit_Log 聯合回答」。
  優點：零 schema 變更；缺點：查詢需跨兩張表關聯，可稽核性較差。

**Claude 傾向方案 A**：此表尚無資料，成本最低，且符合 §18「單表可回答」的原始意圖。

---

### CONFLICT-03 — §24 規定的寫入順序與 CANDIDATE-01 實際順序相反

| Contract §24 規定 | CANDIDATE-01 實際（F-2） |
|---|---|
| Snapshot Card → Write Treatment_Record → Write Card_Usage_Detail → **Update Card_Master** | **DeductSessions（先改 Card_Master）** → Write Treatment_Record → CreateUsageDetail |

Contract 把 `Card_Master` 更新放在**最後**，實際程式放在**最前**。
這不是「尚未實作」，而是**條文與現況直接相反**，且正是 FIND-01 的成因。

**Claude 建議**：採 §24 的順序，並加上 §25 的補償機制。但請 GPT 注意 CONFLICT-04 提到的
「rollback 與不可覆寫歷史」的張力，兩者要一起裁決。

---

### CONFLICT-04 — §25「Rollback Treatment_Record」與 §7 / §12 / §27「歷史不可覆寫」相衝突

§25 要求失敗時「Rollback Treatment_Record / Rollback Card_Usage_Detail / Restore Card_Master」。

但在 Excel 中，「rollback 一筆已寫入的列」實務上就是**刪除該列**，
這與以下條文正面衝突：

- §7「Treatment_Record 必須保存……audit metadata」
- §12「不得修改原 Treatment」
- §27「重大變更必須留下 Who / When / What / Old / New」

若允許刪列，則「某筆治療曾被寫入又被刪除」這件事本身不會留下任何痕跡，
反而**降低**可稽核性。

**Claude 建議：以 VOID 標記取代實體刪除。**
`Treatment_Record` 第 22 欄已有 `Status` 欄位（目前寫入 `"Completed"`），現成可用：

```
失敗時：
  Treatment_Record.Status      = "Voided"
  Card_Usage_Detail            → 新增一筆反向明細（UsageType = "Rollback"）
  Card_Master                  → 回寫 snapshot 前的餘額
  Audit_Log                    → 寫入 Reason = 失敗原因
```

即「補償交易（compensating transaction）」而非「刪除」。
這同時滿足 §25 的資料一致性要求與 §7/§12/§27 的不可覆寫要求，
且 `Card_Usage_Detail` 的 `Version` 欄位（實際存在、目前固定寫 1）正好可用於此。

**Decision Needed（GPT）**：§25 是否改為「Compensating Transaction / VOID marking」而非 rollback-by-deletion。

---

## `[MEDIUM]` FIND-02 — Contract §22 列出的控制項名稱與實際不符，會製造假 Critical

Contract §22 的 `frmTreatment` 控制項範例列出：

```
txtStartTime
txtEndTime
```

**實際控制項名稱為 `txtStart` 與 `txtEnd`**（`modFormBuilder.bas` 第 193、194 行：
`AddText c, "txtStart"` / `AddText c, "txtEnd"`）。

若 Gemini 依 Contract §22 字面進行 §23 的 control 比對，會得到一個**與 Issue 002 完全同型的假陽性 Critical**。
建議 V1.1 直接修正，或改為「以 Claude 提供的 Control Manifest 為準，Contract 不重複列舉」。

---

## `[MEDIUM]` FIND-03 — §8 套卡規格已變更，但無 Package Master 與生效期間

§8 列出的正式規格與專案最初需求文件不同：

| 卡種 | 最初需求文件 | Contract §8 |
|---|---|---|
| 映遊 | 20,000 / 5 瓶 C1 / 20 堂 / 6.5 折 | 16,000 / 4 瓶 C1 / 10 堂 / 9 折 |
| 彩綿 | 24,000 / 6 瓶 C1 / 20 堂 / 6 折 | 24,000 / 6 瓶 C1 / **16 堂** / **8 折** |
| 華夏尊盈 | 60,000 / 4 瓶 C1 / 35 堂 / 6 折 | 60,000 / **5 瓶 C1** / 35 堂 / **7 折** |
| 闔家萬吉 | 100,000 / 10 瓶 C1 / 60 堂 / 6 折 | 100,000 / 10 瓶 C1 / 60 堂 / 6 折 |

§8 說「不得從其他舊版資料自行覆蓋以上正式定義」——**方向正確，但反過來同樣成立**：
2025/12～2026/08 期間購買的歷史套卡是依**舊規格**售出的，
若未來以 §8 的新規格回推歷史卡片的堂數或折扣，會靜默改寫歷史（違反 §28）。

**同時，目前系統中不存在 Package Master 工作表**，§8 的表格只活在文件裡。
`frmCard` 的 `PurchaseAmount` 與 `OriginalSessions` 仍由店長手動輸入，
無任何一致性檢查（例如選「彩綿卡」卻輸入 35 堂，系統不會擋）。

**Claude 建議**：PATCH-03 建立 `Package_Master` 工作表，
欄位比照既有 `Card_Product_Detail` 的版本化慣例，含 `EffectiveFrom` / `EffectiveTo`，
使新舊規格並存、各自適用其期間。**本輪不實作，僅提出。**

---

## `[LOW]` FIND-04 — `CreateUsageDetail` 靜默失敗

`modCardUsage.bas` 第 6 行：`If Len(cardID)=0 Or sessionsUsed<=0 Then Exit Function`
→ 回傳空字串，不拋錯、不留痕跡。無卡客戶屬正常情境，但 `sessionsUsed<=0` 的情況
會靜默跳過，與 §13「Business validation 必須明確存在」的精神不完全一致。優先度低。

---

# Decisions（Claude 對 Contract 各節的立場）

## ACCEPT（無保留同意）

| 節 | 理由 |
|---|---|
| §2 Source of Truth 優先順序 | 這正是解開 21 vs 22 欄混淆的關鍵機制。本次 Review 的 CONFLICT-01/02 也是靠它才發現。 |
| §4 obsolete package 保留於 Archive 而非刪除 | **我先前建議 Arthur「直接刪除」是錯的，GPT 的裁決正確。** 刪掉會失去追溯「為何當時 GPT 看到 21 欄」的證據，反而讓 CONFLICT-02 的成因無法查證。撤回原建議。 |
| §23 Control Mismatch QA Rule | 精準對應已發生的失誤模式，且對稱要求「證明一致時必須標 REJECTED」，不偏袒任一方。 |
| §30 / §31 Evidence Standard 與 Fact / Hypothesis 區分 | 本次 Review 我已全面套用（FIND-01 明確標註「source-level Observed Fact，runtime 未驗證」）。 |
| §32 Version Evidence / SHA-256 manifest | 同意，並承諾下次交付一併提供 `MANIFEST_SHA256.txt`。 |
| §33 Definition of Done | 十級狀態定義清楚，且正確地把「Claude 寫完」與「完成」分開。 |
| §34 Patch Workflow | 同意。特別是禁止「Gemini finding → Claude 立即大改」。 |
| §41 Arthur 最終權限 | 同意，尤其「確認未知 CardID prefix」不得由 AI 代決。 |

## PATCH REQUIRED

- §24 / §25 / §26：需依 CONFLICT-03、CONFLICT-04 調整後才可實作。

## CONFLICT（待 GPT 裁決，見上）

- CONFLICT-01（§17）、CONFLICT-02（§18）、CONFLICT-03（§24）、CONFLICT-04（§25）

## 需修正條文

- §22（FIND-02，控制項名稱錯誤）
- §18（CONFLICT-02，疑似誤植作廢草稿的設計）

---

# Open Issues

| ID | 狀態 | 說明 |
|---|---|---|
| Issue 001 | **建議 CLOSE** | `Treatment_Record` 實際 22 欄，已與 §16 對齊 |
| Issue 002 | **維持 REJECTED** | 但請一併修正 §22 的 `txtStartTime`/`txtEndTime`（FIND-02），否則會再生一次同型假陽性 |
| Issue 003（Atomicity） | **升級為 CRITICAL** | FIND-01 證明它不只是理論風險，而是現在就會造成重複扣堂的具體路徑 |
| Issue 004（映遊 prefix） | OPEN / 待 Arthur | CANDIDATE-01 已設計為主動報錯，不猜測 |
| Issue 005（VBE cache） | 同意維持 UNVERIFIED | Claude 無證據支持或反駁 |
| **Issue 006（新）** | OPEN | CONFLICT-01：Card_Master 缺 CardType，卡種資訊無法稽核 |
| **Issue 007（新）** | OPEN | CONFLICT-02：Card_Usage_Detail 無法回答「扣堂前後餘額」 |
| **Issue 008（新）** | OPEN | FIND-03：無 Package_Master，套卡規格新舊版本未區隔 |

---

# Recommended Action

## 對 Contract V1.1 的具體修改建議

1. §17：改為「附加第 11–15 欄」寫法，並將 `Active` 正名為 `CardStatus`（對齊 artifact）。
2. §18：裁決方案 A 或 B；若採 A，明訂新欄位附加於末尾。
3. §24：修正寫入順序，並明訂「所有 validation（含 shortage resolution）必須在第一次寫入之前完成」。
4. §25：以「Compensating Transaction / VOID marking」取代 rollback-by-deletion，明訂使用
   `Treatment_Record.Status = "Voided"` 與 `Card_Usage_Detail.UsageType = "Rollback"`。
5. §22：修正 `txtStartTime` → `txtStart`、`txtEndTime` → `txtEnd`，或改為引用 Control Manifest。
6. §8：加註「本表適用於 2026-09 起新售套卡；歷史套卡依售出當時規格，不得回溯套用」。
7. §39：新增 Issue 006 / 007 / 008。

## 對 PATCH-02 範圍的建議（優先序）

```
P0  FIND-01 修正（validation 前移）          ← 風險最高、改動最小
P1  CONFLICT-03/04 定案後的 pseudo-transaction
P2  CONFLICT-01/02 的 schema 附加欄位
P3  FIND-02 Control Manifest 交付
P4  FIND-03 Package_Master（可延後至 PATCH-03）
```

**Claude 不會在 GPT 裁決前動任何一行程式碼。**

---

# 對 Arthur 協作路線圖的意見（MD → GitHub → MCP）

**同意分三階段，但建議把 GitHub 提前，理由如下。**

這次專案真正付出代價的問題是「同一個 Claude 產出物有三個版本在流通，
GPT 拿到 A、Gemini 拿到 B，兩邊都沒說錯但結論相反」。
**MD 交接解決不了這個問題**——MD 本身也會有多個版本在流通（Contract §18 誤植作廢草稿的設計，
就是 MD-based 流程下仍然發生的同型錯誤）。

GitHub 直接解決的正是這一類：commit hash 是唯一的，diff 是機械的，
沒有「你看的是哪一份」的空間。而 MCP 解決的是「自動化傳遞」，
那是效率問題，不是正確性問題——目前的瓶頸是正確性，不是效率。

**折衷建議（幾乎零成本，今天就能做）**：在還沒上 GitHub 前，
先落實 Contract §32 的 `MANIFEST_SHA256.txt`。
每份 candidate 附一個 SHA-256 清單，任何 AI 開口前先報 hash。
這能擋掉九成的版本混淆，成本只是每次多附一個文字檔。

MCP 放最後我完全同意——在 schema 和流程都還在變動時做自動化串接，
只會把錯誤傳得更快。

---

# 對這套制度的一點誠實意見

Contract 的方向是對的，這次 Review 就是最好的證明：如果沒有 §2 的 source of truth 排序和
§30 的證據要求，我不會去比對實際工作表，也就不會發現 CONFLICT-01/02 和 FIND-01。
**制度確實抓出了真問題。**

但也想誠實提一件事供 Arthur 衡量：目前每一個 patch 需要
Contract → Patch Spec → Implementation MD → QA MD → Decision MD 五份文件，
對一家單店、三位使用者、以 Excel 為載體的系統來說，儀式相當重。

建議依風險分級，而非一律套用完整流程：

| 變更類型 | 建議流程 |
|---|---|
| 觸及交易、扣堂、金額、schema | 完整流程（現行） |
| 只改 UI 排版、文字、標籤 | Claude 直接改 + Gemini 抽查，不需 Contract 修訂 |
| 修 bug 且不改 schema / signature | Implementation MD + QA MD 即可，略過 Patch Spec |

這是建議，不是反對。最終由 Arthur 決定願意付多少協調成本換多少安全性——
以目前這個系統要管理真實金流與客戶套卡餘額來說，偏嚴謹是合理的選擇。

---

# Next AI Action

| 對象 | 動作 |
|---|---|
| **GPT** | 裁決 CONFLICT-01～04；決定 §18 方案 A/B；修正 §22；產出 Contract V1.1 |
| **Gemini** | 獨立 QA CANDIDATE-01。**請特別複驗 FIND-01**——這是 Claude 自行推導的執行路徑，需要第二雙眼睛。另請驗證 CONFLICT-02 的結論（`Card_Usage_Detail` 確實無 Before/After 欄位） |
| **Arthur** | (1) 執行 CANDIDATE-01 的 `Debug → Compile VBAProject` 並回報；(2) 確認映遊卡 prefix；(3) 裁示 FIND-03 的套卡新舊規格生效日 |
| **Claude** | 待 Contract V1.1 + Gemini QA。**在此之前不動任何程式碼。** |

---

**END OF REVIEW**
