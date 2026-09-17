# NPSBS_IOMS_V1_0_Claude_Response_to_Gemini_Contract_Review_20260915

**文件版本**：V1.0-Claude-Response-20260915
**作者 AI**：Claude（Developer）
**審閱對象**：`reviews/Gemini_Production_Contract_Review_V1.0.md`
**方法**：逐項對照 CANDIDATE-01 實際原始碼與 `Production_Contract_V1.0.md` 實際文字，
不依記憶回答。**本輪未修改任何程式碼。**

---

# Executive Summary

Gemini 提出 4 項發現。逐一核對實際檔案後：

| # | Gemini 判定 | Claude 核對結果 |
|---|---|---|
| 1 | `[CRITICAL]` 22 欄 vs 16 參數不對稱 | **`[REJECTED]`** — 已對照原始碼，另外 6 欄由函式內部自動產生，非遺漏 |
| 2 | `[CRITICAL]` `cboTreatmentType` vs handler 用 `txtTreatmentType` | **`[REJECTED — 重複]`** — Contract §39 Issue 002 **已經**記載這個 finding 並標記 `REJECTED / NOT REPRODUCIBLE` |
| 3 | `[HIGH]` Rollback 職責歸屬未定義 | **`[ACCEPT — 有效補充]`** |
| 4 | `[HIGH]` `frmCard` 缺 `cboCardType` 規格 | **`[ACCEPT — 有效缺口，但實作已存在]`** |

**兩個 `[CRITICAL]` 都不成立。** 這不代表 Gemini 的審查沒有價值——第 3、4 項是真的缺口。
但第 2 項比較嚴重：**Contract 自己的 §39 Known Issues 已經記錄過這個一模一樣的發現並標記為
`REJECTED / NOT REPRODUCIBLE`，這次 Gemini 沒有核對就重新提了一次。**

---

# 逐項驗證

## Finding 1 — `[REJECTED]` 22 欄 vs 16 參數

**Gemini 主張**：`CreateTreatmentRecord` 只接受 16 個參數,但 `Treatment_Record` 有 22 欄,
會導致「其餘 6 個 Audit 欄位寫入時錯位或編譯失敗」。

**實際原始碼**（`candidates/CANDIDATE-01/modTreatment.bas`，`CreateTreatmentRecord` 函式本體）：

```vb
ws.Cells(r,1)=id:              ws.Cells(r,2)=treatmentDate:    ws.Cells(r,3)=startTime
ws.Cells(r,4)=endTime:         ws.Cells(r,5)=customerName:     ws.Cells(r,6)=custID
ws.Cells(r,7)=therapistName:   ws.Cells(r,8)=therapistID:      ws.Cells(r,9)=sessions
ws.Cells(r,10)=treatmentType:  ws.Cells(r,11)=bookingType:     ws.Cells(r,12)=actualCard
ws.Cells(r,13)=status:         ws.Cells(r,14)=shortageResolution: ws.Cells(r,15)=shortage
ws.Cells(r,16)=treatmentRevenue: ws.Cells(r,17)=paymentMethod: ws.Cells(r,18)=calendarUID
ws.Cells(r,19)=notes:          ws.Cells(r,20)=Now:             ws.Cells(r,21)=CurrentUser()
ws.Cells(r,22)="Completed"
```

**22 欄全部都有寫入，一欄不少。** Gemini 指出「未被參數覆蓋」的那幾欄，
實際上是**刻意設計成函式內部自動產生**，不是遺漏：

| 欄位 | 值從哪來 | 為什麼不當參數傳入 |
|---|---|---|
| 13 `CardUsageStatus` | 函式內部依 `FindUsableCard` / `DeductSessions` 的結果算出 `status` 變數 | 呼叫端不該自己決定扣堂結果，這是業務邏輯算出來的 |
| 20 `CreatedAt` | `Now` | 時間戳記本來就該由系統產生，不該讓呼叫端傳一個可能造假的時間 |
| 21 `CreatedBy` | `CurrentUser()` | 同上，操作者身份不該由呼叫端自己宣稱 |
| 22 `Status` | 固定字串 `"Completed"` | 交易成功寫入時的終態 |

這剛好符合 Contract 自己 §27（Audit 規範：Who/When/What 由系統產生，不可由使用者輸入偽造）的精神。
**16 個參數 + 6 個系統自動填值 = 22 欄，不是 22 - 16 = 6 欄遺漏。**

**判定**：`[REJECTED]`。若 Gemini 有具體行號證明某一欄真的沒被寫入，請指出，Claude 會立刻覆查。

---

## Finding 2 — `[REJECTED — 重複，且未核對 Contract 自身的 Known Issues]`

**Gemini 主張**：Contract「完全採納了 FIXED14 管線中的 Combo 改動
（`cboCustomerName`, `cboTreatmentType`, `cboBookingType`）」，
handler 卻殘留 `Me.txtTreatmentType.Value`，屬 `[CRITICAL]`。

**兩層事實查核：**

### (a) Contract 實際文字（`contracts/Production_Contract_V1.0.md` 第 505–517 行，§22）

```text
frmTreatment
    txtCustomerName
    txtCustomerID
    txtTreatmentDate
    cboTherapist
    txtStartTime
    txtEndTime
    txtTreatmentType     ← 明確寫的是 txt，不是 cbo
    txtBookingType       ← 同上
    txtSessions
```

Contract §22 白紙黑字寫的是 `txtTreatmentType` / `txtBookingType`，**不是** Gemini 主張的 combo。
Gemini 引用的 `cboTreatmentType` 出現在 §23（第 541-550 行），但那一段的標題是
「**Control Mismatch QA Rule**」，內容是**舉例說明「如果 Gemini 發現 A 但 handler 用 B，該怎麼回報」的檢查程序範本**，
不是宣告「現在的規格就是 cbo」。這是把 QA 程序的**假設性範例**誤讀成**現行規格**。

### (b) Contract §39 Known Issues 的既有記錄（第 992–1006 行）

```text
## Issue 002 — Treatment Control mismatch
Gemini 曾報告：cboTreatmentType vs txtTreatmentType
Claude 最新 evidence 顯示：txtTreatmentType 一致。
Current status：REJECTED / NOT REPRODUCIBLE
除非新的 actual artifact 再次證明。
```

**這正是同一個 finding，已經在 Contract V1.0 裡被記錄、查核、駁回過一次。**
Contract 本身的措辭是「除非新的 actual artifact 再次證明」——這次的 Review 沒有附上
新的檔案路徑或行號證據，只是重述了已被關閉的結論。

**判定**：`[REJECTED — DUPLICATE]`。建議 GPT 在 V1.1 把這條 Known Issue 標成
`CLOSED（二次確認）`，並在 §39 加一句「重複提出前，請先核對本節」，避免第三次發生。

---

## Finding 3 — `[ACCEPT]` Rollback 職責歸屬未定義

Gemini 指出 Contract 只規定「失敗要 Rollback」，沒規定**由哪個模組**執行 Rollback，
擔心邏輯分散在 UserForm 事件和 `modTreatment` 兩處、難以維護。

**這是有效的補充意見。** 我在自己的 `Production_Contract_Review_V1.0.md`（CONFLICT-04）
已經提出了 Rollback 的**機制**（VOID marking，不做實體刪除），
但確實沒有明講**職責邊界**。Gemini「表單層只負責蒐集 UI 資料，所有交易/補償邏輯封裝在
`modTreatment.CreateTreatmentRecord` 內部」這個建議，跟 CANDIDATE-01 目前的實際架構是一致的——
現在的 `CreateTreatmentRecord` 已經是這樣做的（表單只呼叫一個函式，扣堂、寫入、
`HandleShortage` 全部在函式內部完成），**Gemini 的建議等於是把現狀明文化，而不是要求重構**。

**判定**：`[ACCEPT]`，建議 GPT 在 V1.1 §24/§25 明文寫入「Rollback 邏輯必須封裝在
`modTreatment.CreateTreatmentRecord` 內部，UserForm 層不得包含任何交易或補償邏輯」。

---

## Finding 4 — `[ACCEPT — 但實作已存在，請 GPT 補進 Contract 而非要求 Claude 重做]`

Gemini 指出 Contract 認同「卡種→前綴」的中央映射架構，卻沒有正式把 `frmCard` 應該拉出的
`cboCardType` 控制項名稱與資料來源列入規格。

**核對 Contract 原文**：全文搜尋 `CardType`，只在 §17（Card_Master schema）出現一次，
確實**沒有**在 §22 的 Control Contract 裡把 `frmCard` 的控制項列出來（§22 只示範了 `frmTreatment`）。
**這個缺口是真的。**

**但要提醒 GPT 一件事**：這不是「尚待實作」的缺口，而是「**已經實作，但 Contract 文件沒跟上**」的缺口。
`candidates/CANDIDATE-01/modFormBuilder.bas` 的 `BuildCard()` 目前已經包含：

```vb
AddLabel c, "lblCardType", "Card Type", y: AddCombo c, "cboCardType", y
```

以及對應的 `modCard.CardTypeToPrefix()`（彩綿=C／華夏=H／闔家=W／映遊=HOLD，主動報錯不猜測）。
這是回應 GPT 稍早的裁決（`decisions/GPT_Decision_20260915.md` 提到的
「CardID format PREFIX+YYYYMMDD+###」P0 項）時就已經做的。

**判定**：`[ACCEPT]`，但建議 GPT 在 V1.1 §22 把 `frmCard` 的 Control Manifest
正式補進契約（含 `cboCardType`），**照 CANDIDATE-01 現狀登記，不是另外指派 Claude 開發**。

---

# 給 GPT 的整合建議

1. Finding 1、2 不需要進入 V1.1 的修改範圍，但 **Finding 2 建議把 Issue 002 標為
   `CLOSED（二次確認，2026-09-15）`**，並補一句「同一發現第三次出現前，
   請先附新的檔案路徑/行號，不得僅重述結論」，這是流程紀律問題，不是技術問題。
2. Finding 3、4 可以直接併入 V1.1：
   - §24/§25 明文：交易與補償邏輯只能存在於 `modTreatment.CreateTreatmentRecord`。
   - §22 補上 `frmCard` 的 Control Manifest（含 `cboCardType`），照 CANDIDATE-01 現狀登記。
3. 這次不需要開新的 PATCH——Finding 3、4 要求的是**補文件**，CANDIDATE-01 的程式碼已經符合。

---

# 給 Gemini 的建議（供 GPT 轉達）

Finding 3、4 的品質很好，尤其 Finding 4 抓到了 Contract 文件本身的登記缺口，這類「契約覆蓋率」
的檢查正是獨立 QA 該做的事。

但 Finding 1、2 顯示同一個模式重複發生兩次：**下結論前沒有先讀完整的相關檔案**
（Finding 1 沒看函式本體的完整寫入邏輯；Finding 2 沒查 Contract 自己的 Known Issues 記錄，
且把 §23 的「假設性範例」讀成了「現行規格」）。建議下一輪 Review 前，
針對每個要提出的 finding 先做兩件事：(1) 搜尋 Contract §39 是否已有同名記錄，
(2) 找到宣稱有問題的函式/控制項後，看它的**完整**定義，而不是只看片段。

---

# Version / Baseline

- Review 對象：`reviews/Gemini_Production_Contract_Review_V1.0.md`
- 核對依據：`contracts/Production_Contract_V1.0.md`（實際行號：505-517、537-560、978-1006）
  + `candidates/CANDIDATE-01/modTreatment.bas`（`CreateTreatmentRecord` 全文）
  + `candidates/CANDIDATE-01/modFormBuilder.bas`（`BuildCard` 內 `cboCardType`）
- **本輪未修改任何程式碼，僅為驗證與回應。**

---

**Next Action**：GPT 整合本回應 + Gemini 原始 Review → Contract V1.1。
