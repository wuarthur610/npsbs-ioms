# NPSBS IOMS V1.0 — Claude Patch Implementation

**文件名稱**：NPSBS_IOMS_V1_0_Claude_Patch_Implementation_20260911.md
**文件版本**：V1.0-Claude-20260911
**日期**：2026-09-11
**作者 AI**：Claude
**文件用途**：回應 GPT `Three_AI_Collaboration_Handoff_20260911` 第 29 節的直接指示，
提供實際 candidate artifact、釐清 21 vs 22 column 版本衝突、交付 PATCH-01 實作。
**所依據的 baseline**：`NPSBS_IOMS_V1_0_Production_VBA_UserForm_FINAL_FIXED14.zip`
（內含 17 個 `.bas`，檔頭標示 FIXED13）

---

# 1. 給 Arthur 的白話結論

- GPT 說「Claude 的包還是 21 欄、還在用 `CreateTreatmentTransaction`」——**這是在看一份已經作廢的舊草稿**。
  那是我在你上傳真正的 `.xlsm` 之前、憑文字描述從零寫的 MVP，我當時就說過作廢了。
  **請把 `NPSBS_IOMS_V1_MainSystem_Package.zip` 整個刪掉**，不要再給任何 AI 看，混淆就是從它來的。
- Gemini 說的那個「Critical：表單建了 `cboTreatmentType`、程式卻讀 `txtTreatmentType`」——
  **我實際檢查過原始碼，這個問題不存在**。FIXED14 建的就是 `txtTreatmentType`，讀的也是它，兩邊一致。
- 但我在檢查過程中**找到兩個真的問題**，都已經修好了：
  1. FIXED14 裡有一句錯誤訊息（找不到待補扣紀錄時）忘了翻譯，會跳出英文代碼給店長看。
  2. 我上一版的 `frmShortage` 沒防呆，同一筆治療可以被重複處理兩次、產生兩筆待補扣紀錄。
- 另外，GPT 指出的「卡號前綴要依卡種決定」我也做了。**映遊卡的前綴你還沒確認，所以我沒有亂猜——
  系統會直接跳出訊息請你先確認**，而不是偷偷給它一個 `C`。
- 這次改動只碰 3 個模組，其餘 14 個跟 FIXED14 一模一樣（連空白都沒動）。

---

# 2. Technical Decision

1. **以 FIXED13/14 官方 `.bas` 為 rebase 基底**，不是以我之前的交付物為基底。
   上一版我的 `modFormBuilder.bas` 是從 `.xlsm` 還原出來的 VBE-normalized 版本
   （`ctl.Left` 被正規化成 `ctl.left`），雖然 VBA 不分大小寫、功能等價，但為了符合
   GPT 第 15 節「Source Integrity」原則，本次全部重做成以官方 `.bas` 逐字為基底。
2. **遵守 ASCII-safe 原始碼規範**。上一版我在 `modFormBuilder.bas` 寫了中文註解，
   違反第 2.1 節的既定架構規範，本次已全部移除，並以工具驗證 17 個檔案皆為純 ASCII。
   新增的中文字串一律走 `modText.UText` / `MsgText` 的 `ChrW` 慣例。
3. **不進行架構改寫**。UserForm builder 架構、`CreateTreatmentRecord` 簽章、
   22 欄 schema 全部維持 FIXED14 原樣。
4. **不猜測未確認事項**。映遊卡 prefix 未確認 → 程式主動 raise error，不填預設值。

---

# 3. Confirmed Facts（有實際檔案證據）

## 3.1 [CONFLICT 已解決] 21-column vs 22-column

GPT 第 14 節列為「目前最大的版本衝突」。**結論：不是版本衝突，是兩份不同檔案。**

| 檔案 | 內容 | 狀態 |
|---|---|---|
| `NPSBS_IOMS_V1_MainSystem_Package.zip` | 21 欄 A~U、`CreateTreatmentTransaction`、`frmDailyTreatment`、`modUI`、`modShortage` skeleton | **已作廢**。Claude 在拿到真實 `.xlsm` 前，純憑文字需求描述寫的 MVP 草稿。Claude 當時已明確告知 Arthur 作廢。 |
| `NPSBS_IOMS_RuntimeBinding_Update.zip` | 只有 `modCard.bas` + `modFormBuilder.bas`，22 欄、`CreateTreatmentRecord` | 上一版交付，方向正確但有第 4 節列出的瑕疵 |
| **`NPSBS_IOMS_V1_0_CANDIDATE-01.zip`（本次）** | 17 個完整 `.bas`，22 欄、`CreateTreatmentRecord` | **本次唯一有效 candidate** |

Gemini 看到 22 欄、GPT 看到 21 欄，兩邊都沒說錯，只是看了不同檔案。
**建議 Arthur 直接刪除 `NPSBS_IOMS_V1_MainSystem_Package.zip`，杜絕再次混淆。**

## 3.2 [已駁回] Gemini 的 Critical：UserForm control / handler mismatch

GPT 第 16 節引述 Gemini 指出 `BuildTreatment` 建立 `cboTreatmentType` / `cboBookingType`
但 handler 用 `Me.txtTreatmentType.Value`。

**實際檔案證據**（`FIXED13/modFormBuilder.bas` 第 196、197、225 行，逐字引用）：

```vb
196:  AddLabel c, "lblType", "Treatment Type", y: AddText c, "txtTreatmentType", y: y = y + 32
197:  AddLabel c, "lblBooking", "Booking Type", y: AddText c, "txtBookingType", y: y = y + 32
225:  ... CDbl(Me.txtSessions.Value), Me.txtTreatmentType.Value, Me.txtBookingType.Value, ...
```

建立的是 `AddText`（TextBox，名為 `txtTreatmentType`），handler 讀的也是 `txtTreatmentType`。
**整個 FIXED13 原始碼中不存在 `cboTreatmentType` 或 `cboBookingType` 這兩個識別字。**

Claude 另外對全部 8 個 `Build*` Sub 做了程式化的靜態交叉檢查
（比對每個 Sub 建立的控制項名稱 vs 注入程式碼中所有 `Me.<name>` 引用），
結果：**8 個表單全數通過，沒有任何 handler 引用不存在的控制項。**

→ 建議標記為 `[REJECTED - NOT REPRODUCIBLE]`。
但 **GPT 第 16 節提出的「UI Control → Handler → Signature → Columns 必須一致」這條原則本身完全正確**，
Claude 已將它實作成可重複執行的自動檢查（見第 7 節）。

## 3.3 Production Schema（本 candidate 實際使用，與 GPT 第 9 節一致）

`Treatment_Record` 22 欄，來源為 `modSystemCore.EnsureProductionSheets` 的表頭陣列：

```
1 TreatmentRecordID   2 TreatmentDate      3 TreatmentStart_Taipei  4 TreatmentEnd_Taipei
5 CustomerName        6 CustomerID         7 Therapist              8 TherapistID
9 Sessions           10 TreatmentType      11 BookingType           12 CardID
13 CardUsageStatus   14 ShortageResolution 15 ShortageSessions      16 TreatmentRevenue
17 PaymentMethod     18 CalendarUID        19 Notes                 20 CreatedAt
21 CreatedBy         22 Status
```

## 3.4 實際函式簽章（本 candidate）

```vb
' modTreatment.bas — 未修改，維持 FIXED14 原樣
Public Function CreateTreatmentRecord(ByVal treatmentDate As Date, ByVal startTime As Date, _
    ByVal endTime As Date, ByVal customerName As String, ByVal therapistName As String, _
    ByVal therapistID As String, ByVal sessions As Double, ByVal treatmentType As String, _
    ByVal bookingType As String, Optional ByVal cardID As String = "", _
    Optional ByVal paymentMethod As String = "", Optional ByVal treatmentRevenue As Double = 0, _
    Optional ByVal calendarUID As String = "", Optional ByVal notes As String = "", _
    Optional ByVal shortageResolution As String = "", Optional ByVal shortageAmount As Double = 0) As String
' → 9 個必要參數 + 7 個選擇性參數。名稱是 CreateTreatmentRecord，
'   本 candidate 中不存在 CreateTreatmentTransaction。

' modCard.bas — PATCH-01 新增
Public Function CardTypeToPrefix(ByVal cardType As String) As String
Public Function CreateCard(ByVal customerID As String, ByVal cardType As String, _
    ByVal purchaseDate As Date, ByVal purchaseAmount As Double, ByVal originalSessions As Double, _
    Optional ByVal paymentMethod As String = "", Optional ByVal cardID As String = "", _
    Optional ByVal notes As String = "") As String
```

---

# 4. Findings / Issues

| # | 等級 | 問題 | 來源 | 狀態 |
|---|---|---|---|---|
| F1 | `[RESOLVED]` | 21 vs 22 column「版本衝突」 | GPT §14 | 非衝突，是已作廢草稿；見 3.1 |
| F2 | `[REJECTED]` | `cboTreatmentType` / handler mismatch | Gemini via GPT §16 | 原始碼證明不存在；見 3.2 |
| F3 | `[HIGH]` **新發現** | `modCardUsage` 第 38 行呼叫 `MsgText("PENDING_NOT_FOUND")`，但 `modText` **從未定義這個 key**。`MsgText` 的 `Case Else` 會原樣回傳 key，店長會看到英文字串 `PENDING_NOT_FOUND` 而不是中文訊息。 | Claude 靜態檢查 | **PATCH-01 已修**（`modText` 新增該 key） |
| F4 | `[HIGH]` **自我檢出** | Claude 上一版 `frmShortage` 沒有防止重複處理：同一 `TreatmentRecordID` 可被處理兩次，產生兩筆 `Pending_Shortage`，違反「不得重複扣」原則 | Claude 自我審查 | **PATCH-01 已修**（加入 col 14 已填值 / shortageQty<=0 兩道 guard） |
| F5 | `[MEDIUM]` **自我檢出** | Claude 上一版 `modFormBuilder.bas` 含中文註解，違反 §2.1 ASCII-safe 規範；且以 VBE-normalized 還原碼為基底而非官方 `.bas` | Claude 自我審查 | **PATCH-01 已修**（重新 rebase + ASCII 驗證） |
| F6 | `[RESOLVED]` | `GenerateCardID` 固定 `"C"` 前綴，不符卡種規則 | GPT §17/§20 | **PATCH-01 已修**；見第 5.1 節 |
| F7 | `[OPEN]` | **交易 atomicity**：`CreateTreatmentRecord` 先扣卡後寫紀錄，中途失敗會留下「卡已扣、無治療紀錄」 | Gemini via GPT §17 | **本次未動**。這需要改 `modTreatment` 核心流程，屬架構層級，依 §28「Contract 未鎖定前不做大規模 rewrite」暫緩，等 GPT Production Contract。 |
| F8 | `[BLOCKED]` | 映遊卡 CardID prefix 未確認 | GPT §5 | 需 Arthur 確認。程式已設計成「主動報錯」而非猜測。 |
| F9 | `[OPEN]` | 套卡規格（價格/堂數/折扣）目前未在程式中集中定義，`frmCard` 的金額與堂數仍由店長手動輸入 | Claude | 建議納入 Production Contract 後再實作 Package Master |

---

# 5. Proposed Changes（PATCH-01 實際內容）

**只改 3 個模組，其餘 14 個與 FIXED13 逐字節相同。**

| 模組 | 變更 | 行數 |
|---|---|---|
| `modCard.bas` | 新增 `CardTypeToPrefix`、`CreateCard`、`CardIDExists`、`CustomerExists`。**原有 6 個函式完全未動。** | +82 / -0 |
| `modFormBuilder.bas` | 6 個 `Build*` Sub 補 runtime binding；移除 6 個佔位 `Save*Form`；`BuildCard` 增加 `cboCardType`。**`BuildMainMenu` / `BuildTreatment` / 所有 helper 完全未動。** | +143 / -40 |
| `modText.bas` | 新增 8 個 MsgText key（含修補 F3）+ 4 個卡種名稱函式。**原有內容未動。** | +25 / -0 |

## 5.1 CardID prefix 中央化（回應 GPT §20）

```vb
Public Function CardTypeToPrefix(ByVal cardType As String) As String
    Dim t As String
    t = Trim$(cardType)
    If t = TW_CardCaimian() Then          ' 彩綿卡
        CardTypeToPrefix = "C"
    ElseIf t = TW_CardHuaxia() Then       ' 華夏尊盈卡
        CardTypeToPrefix = "H"
    ElseIf t = TW_CardHejia() Then        ' 闔家萬吉卡
        CardTypeToPrefix = "W"
    ElseIf t = TW_CardYingyou() Then      ' 映遊卡
        Err.Raise vbObjectError + 1260, "IOMS", MsgText("CARD_PREFIX_UNCONFIRMED")
    Else
        Err.Raise vbObjectError + 1261, "IOMS", MsgText("CARD_TYPE_UNKNOWN")
    End If
End Function
```

符合 §20「prefix 由中央 Card 模組管理，不散落在 UserForm」，
以及 §25「映遊 CardID prefix 不得自行猜測」——選擇主動報錯而非填預設值。
日流水號沿用 FIXED14 既有的 `GenerateCardID(prefix, purchaseDate)`，未重寫。

## 5.2 `CreateCard` 的驗證（回應 GPT §18 的 10 點檢查表）

| GPT 要求 | 實作狀態 |
|---|---|
| CardType | 已加為第 2 個必要參數 |
| Card prefix | `CardTypeToPrefix` |
| 同日流水號 | 沿用既有 `GenerateCardID` |
| PaymentMethod 正式枚舉 | 用既有 `TW_Cash()` / `TW_Bank()`，非自由文字 |
| PurchaseAmount >= 0 | 已驗證 |
| OriginalSessions > 0 | 已驗證 |
| CustomerID 是否存在 | `CustomerExists()` 已驗證 |
| CardID 不得重複 | `CardIDExists()` 已驗證 |
| Audit_Log | 已呼叫 `WriteAuditLog` |
| Card_Product_Detail 流程 | **未實作**，留待 Contract（見 F9） |

---

# 6. Impact

**風險低**：
- 14 個模組零變更 → 對 `frmMainMenu`、`frmTreatment`、既有交易流程**零迴歸風險**。
- `modCard` / `modText` 是純新增，沒有修改任何既有函式的行為或簽章。
- `modFormBuilder` 的變更集中在 6 個原本只有佔位訊息的 Sub，
  等於「從沒有功能」變成「有功能」，不會讓任何原本可動的東西壞掉。

**需注意**：
- `BuildProductionUserForms` 會**移除並重建全部 8 個表單**，包含已驗證的
  `frmMainMenu` / `frmTreatment`。它們的 builder 程式碼未變更，重建後應與現況相同，
  但這是唯一會碰到既有已驗證成果的動作，**建議先備份 `.xlsm`**。
- `frmCard` 新增 `cboCardType` 控制項 → 若 Gemini 有既有的 control name 清單，需同步更新。

---

# 7. Files Examined

| 檔案 | 用途 | 方法 |
|---|---|---|
| `NPSBS_IOMS_V1_0.xlsm` → `xl/vbaProject.bin` | 取得 Excel 中實際載入的 VBA | Claude 自寫的純 Python CFBF + MS-OVBA 解壓縮器（沙盒無網路，無法安裝 oletools），還原 27 個 code module |
| `NPSBS_IOMS_V1_0_Production_VBA_UserForm_FINAL_FIXED14.zip` → 17 個 `.bas` | 官方原始碼基底 | 逐行 diff vs 上述還原結果 → **17 個模組邏輯 100% 相同**，差異僅 VBE 對大小寫/空白的正規化 |
| `.xlsm` 各工作表表頭 | 驗證欄位順序 | `openpyxl` 讀取 `Card_Master` / `Card_Usage_Detail` / `Card_Product_Detail` / `Inventory` / `Pending_Shortage` / `Customer_Master` / `Treatment_Record` 實際第 1 列 |
| `NPSBS_IOMS_V1_0_GPT_Latest_Progress_Handoff_20260911.md` | GPT 前一份交接 | 全文閱讀 |
| `NPSBS_IOMS_V1_0_Three_AI_Collaboration_Handoff_20260911.md` | GPT 協作規範 | 全文閱讀，本文件依其 §24 格式撰寫 |

## 本次執行的自動化驗證（可重現）

1. **ASCII-safe 檢查**：17 個 `.bas` 全部通過（以 `errors='strict'` 寫入 ASCII 驗證）。
2. **控制項 vs handler 交叉檢查**：解析每個 `Build*` Sub 建立的控制項名稱，
   與注入程式碼中所有 `Me.<name>` 引用比對 → 8 個表單全數通過。
3. **呼叫參數數量 vs 簽章檢查**：還原注入字串的實際 VBA 內容（正確處理
   `& DQ() &` / `& vbCrLf &` 串接），比對被呼叫函式的必要/選擇性參數數量 →
   `CreateTreatmentRecord`(16/9-16)、`CreateCustomer`(2/1-2)、`CreateCard`(7/5-8)、
   `SaveCardProductVersion`(5/5-6)、`HandleShortage`(7/7)、`ProcessCountVariance`(4/4-5)
   全數在範圍內。
4. **MsgText key 完整性檢查**：所有被引用的 key 都有定義（此項檢查找出 F3）。
5. **TW_* 函式完整性檢查**：全數有定義。

**限制（誠實聲明）**：Claude 的執行環境沒有 Windows / Excel / VBA 編譯器，
以上皆為靜態分析，**不等於 `Debug → Compile VBAProject` 通過**。
真正的 compile 驗證必須由 Arthur 在 Excel 中執行。

---

# 8. Version / Baseline

```
FIXED14 (= FIXED13 .bas ×17, Arthur 已驗證 Compile PASS)
    ↓  PATCH-01 (modCard / modFormBuilder / modText，其餘 14 個逐字節不變)
CANDIDATE-01   ← 本次交付
    ↓  待 Gemini QA
QA-01
    ↓  待 Arthur UAT
UAT-01
```

**交付檔名**：`NPSBS_IOMS_V1_0_CANDIDATE-01_20260911.zip`
**內含**：`VBA_Modules/`（17 個 `.bas`）+ 本文件
**未覆蓋 FIXED14**，符合 §23。

---

# 9. Remaining Questions

### Decision Needed — Arthur

1. **映遊卡的 CardID prefix 是什麼？**（F8）
   - 現況：程式會報錯擋住，無法用映遊卡種自動產生卡號（可手動輸入卡號繞過）
   - 建議：確認後 Claude 一行改完
2. **是否同意刪除 `NPSBS_IOMS_V1_MainSystem_Package.zip`？**
   - 它是所有版本混淆的來源，留著只會讓下一輪再錯一次

### Decision Needed — GPT

3. **F7 交易 atomicity 要不要納入 PATCH-02？**
   Claude 的看法：Excel VBA 沒有真正的 transaction，可行做法是
   「先驗證全部 → 先寫 Treatment_Record → 再扣卡 → 失敗則刪除剛寫入的列 + 還原卡片餘額 + 寫 Audit」。
   但這會動到 `modTreatment` 核心，依 §28 需要 GPT 先鎖 Contract。**Claude 不會自行動手。**
4. **`frmShortage` 的定位是否如 Claude 所理解？**
   Claude 的實作前提是：它只處理「當初 `ShortageResolution` 留空」的既有紀錄，
   **不重新扣堂**（扣堂只在 `CreateTreatmentRecord` 發生一次），只寫 col 14 +
   視情況建立 `Pending_Shortage`。若 GPT 對它的定位不同，請明確指正。
5. **套卡規格（價格/堂數/折扣）是否要建 Package Master 工作表？**（F9）

### For Gemini

6. 請以 `CANDIDATE-01` 為 QA source，**不要再用任何舊 package**。
7. 請特別複驗 F2（Claude 主張 Gemini 的 Critical 不成立）。如果 Gemini 有具體檔案證據
   證明 `cboTreatmentType` 確實存在於某個版本，請指出**檔名 + 行號**，Claude 會立刻重驗。
8. 請複驗 F4 的 guard 是否足夠（Claude 自己找出的問題，自己驗可能有盲點）。

---

# 10. Next Action

1. **Arthur**：備份現有 `.xlsm` → 匯入 CANDIDATE-01 的 17 個 `.bas` →
   `Debug → Compile VBAProject` → 回報結果（成功或完整錯誤訊息＋行號）。
   **先不要跑 UAT**，等 compile 確認。
2. **Gemini**：對 CANDIDATE-01 做獨立 QA，輸出
   `NPSBS_IOMS_V1_0_Gemini_QA_Audit_20260911.md`。
3. **GPT**：建立 `NPSBS_IOMS_V1_0_Production_Contract.md`，特別是 F7 的
   atomicity 規範與 F4 的 shortage 生命週期。
4. **Claude**：等 Contract + Gemini QA，不自行擴充功能。

---

# 11. Handoff MD

本文件即為交接 MD。

**Claude 對 §29 八點指示的逐項回應：**

| GPT 指示 | 回應 |
|---|---|
| 1. 提供實際 candidate `.xlsm` / `.bas` | 已提供 17 個 `.bas`。**未提供 `.xlsm`**——Claude 環境無法產生含編譯後 VBA 的 `.xlsm`（`vbaProject.bin` 只有 Excel 能生成），這是環境限制，非疏漏。 |
| 2. 明確指出版本與日期 | CANDIDATE-01 / 2026-09-11 / rebase 自 FIXED13 `.bas` |
| 3. 列出與 FIXED14 相比修改的 modules | 3 個：`modCard`、`modFormBuilder`、`modText`；其餘 14 個逐字節相同（見第 5 節） |
| 4. 提供實際 `Treatment_Record` schema | 22 欄，見 3.3，與 GPT §9 完全一致 |
| 5. 提供實際函式簽章 | 見 3.4。**是 `CreateTreatmentRecord`，不是 `CreateTreatmentTransaction`** |
| 6. 不要只提供 handoff 描述 | 已附實際 `.bas` |
| 7. 若 22-column candidate 確實存在，提供 source artifact | 已提供。它一直存在，只是 GPT 手上拿到的是另一份已作廢的包 |
| 8. Contract 未鎖定前不進行大規模架構改寫 | 已遵守。F7 atomicity 明確標為 `[OPEN]` 未動。 |
