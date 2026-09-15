# 🔎 reviews/Gemini_QA_Audit_20260915.md
**文件名稱**：NPSBS IOMS V1.0 — CANDIDATE-01 獨立代碼審計與原始碼白盒測試報告
**文件版本**：V1.0-Gemini-QA-CANDIDATE-01-20260915
**當前日期**：2026-09-15
**作者 AI**：Gemini (Independent QA / Code Reviewer)
**文件用途**：針對 GitHub 共享庫中 `candidates/CANDIDATE-01/` 的 17 個實體模組進行頂層合約符合性稽核
**審計對象**：`CANDIDATE-01` 17 Modules (Pure ASCII .bas)
**所依據之至高契約**：`V1.0-GPT-20260911` (三方 AI 共同工作交接與版本控管規範)

---

## 一、 審計官認知基準宣告 (Auditor Baseline Statement)
⚖️ **Gemini 宣告**：本審計報告嚴格遵循「Source of Truth 優先順序原則」。所有探測到的衝突與漏洞，皆以 `candidates/CANDIDATE-01/` 目錄下實際存在的 `.bas` 代碼行為唯一物理證據，不採信任何口頭或文字 Claim。本階段全面凍結手動修改，專注於輸出規格缺陷。

---

## 二、 核心 Schema 與控制項強一致性審計 (Control & Field Mapping)

### 1. `[CRITICAL]` `frmTreatment` 執行期控制項成員跑位錯誤 (Control Mismatch)
* **檔案路徑**：`candidates/CANDIDATE-01/modFormBuilder.bas`
* **模組程序**：`Private Sub BuildTreatment()`
* **定位行號**：第 164 行、第 165 行 與 第 185~192 行
* **代碼事實**：
  * 在第 164-165 行，Builder 物理上添加控制項：
    ```vba
    AddCombo c, "cboTreatmentType", y
    AddCombo c, "cboBookingType", y
    ```
  * 在第 185 行注入的 `SaveTreatmentForm` 核心程序中，底層實際呼叫 `CreateTreatmentRecord` 的字串卻硬編碼寫成：
    ```vba
    Me.txtTreatmentType.Value
    Me.txtBookingType.Value
    ```
* **審計判定**：此項屬於**嚴重的執行期死鎖缺陷**。由於 `txtTreatmentType` 工作表外殼已改為 `Combo`，執行期一經調用必引發 `Object Required` (錯誤 424) 崩潰。
* **修改令**：必須將注入字串精準替換為 `Me.cboTreatmentType.Value` 與 `Me.cboBookingType.Value`。

### 2. `[CRITICAL]` `frmCard` 前綴硬編碼造成編碼系統失控
* **檔案路徑**：`candidates/CANDIDATE-01/modFormBuilder.bas`
* **模組程序**：`Private Sub BuildCard()`
* **定位行號**：第 205 行 至 第 225 行
* **代碼事實**：
  * 整個 `BuildCard` 區塊完全缺失 `cboCardType` 下拉選單控制項的宣告與物理建立。
  * 在第 218 行注入的 `SaveCardForm` 字串中，調用 `CreateCard` 時前綴參數被硬編碼（Hard-coded）鎖死為固定字串 `"C"`：
    ```vba
    rid = CreateCard(Me.txtCustomerID.Value, CDate(Me.txtPurchaseDate.Value), ..., "C")
    ```
* **審計判定**：這直接違反了合約第 5 節與第 20 節關於「卡種前綴中央映射」的強制規範。會導致華夏卡 (`H`) 與闔家卡 (`W`) 的帳目數據被全數洗成彩綿卡格式。
* **修改令**：必須在 `BuildCard` 補齊 `AddCombo c, "cboCardType", y`，且事件注入端必須調用中央映射，禁止硬編碼常數。

---

## 三、 核心商業交易與安全邊界審計 (Transaction & Balance Guard)

### 1. `[HIGH]` `CreateTreatmentRecord` 交易原子性（Atomicity）完全真空
* **檔案路徑**：`candidates/CANDIDATE-01/modTreatment.bas`
* **模組程序**：`Public Function CreateTreatmentRecord(...)`
* **定位行號**：第 45 行 至 第 85 行
* **代碼事實**：
  * 程序在第 52 行呼叫 `DeductSessions` 成功更新 `Card_Master` 工作表單。
  * 隨後向下展延寫入 `Treatment_Record` 工作表。經查，其下方對應的 `EH:`（Error Handler）異常處理區塊中，**完全沒有任何 Rollback 數據恢復邏輯**。
* **審計判定**：此處存在高度資金風險。若寫入 `Treatment_Record` 失敗（例如超過 Excel 最大列數限制），客戶的套卡餘額已被扣除，但消費明細卻完全為空，產生資產不對稱。
* **修改令**：必須在扣堂前建立記憶體快照 `Dim beforeRemaining As Double`，並在 `EH:` 異常區塊中執行逆向還原（Rollback）物理覆寫。

### 2. `[HIGH]` Balance Guard 餘額零底線硬阻斷缺失
* **檔案路徑**：`candidates/CANDIDATE-01/modCard.bas`
* **模組程序**：`Public Function DeductSessions(...)`
* **定位行號**：第 88 行 至 第 112 行
* **代碼事實**：
  * 代碼仰賴外層邏輯保護，在物理變更儲存格 `Cells(i, 7).Value` 前，**未設置顯式的零底線邏輯閘防禦**。
* **審計判定**：由於專案模組大量封裝了 `On Error Resume Next`，一旦外層傳入異常參數，計算過程被吞噬，系統將會直接突破零底線限制，將套卡餘額洗成**負數（Negative Balance）**。
* **修改令**：必須在物理減算行上一列，強制卡入 `If beforeValue < requestedValue Then DeductSessions = False: Exit Function` 閘門。

---

## 四、 尚未確認事項清單 (Unverified Items)
* `[UNVERIFIED]`：因缺乏真實 Excel 環境的二進位執行驗證，`modInventory.bas` 與 `modProduct.bas` 內部對應表單的動態反向寫入邏輯目前列為 `[UNVERIFIED]`，待本輪合約調和後再行測試。
* `[UNVERIFIED]`：映遊卡之正式 CardID Prefix 常數（目前代碼中完全予以安全隔離，符合合約第 25 節禁令）。

---

## 五、 結論與下一步行動倡議 (QA Gate Conclusion)

綜上所述，`candidates/CANDIDATE-01/` 雖然在全英文架構與自動化建構（Builder Shell）上實現了穩定的灰色編譯，但由於**執行期控制項跑位與偽交易回滾（Rollback）機制真空**，本代碼包無法投入生產環境。

**Gemini QA 最終判定：【REJECT（拒絕放行）】**。

建議 Product Owner（亞瑟）將此文件匯出，直接指令 **GPT 夥伴** 進入 Step 2 發布對應的 `PATCH-01` 規格書，由 Claude 進行精準阻斷性修正！

---
**Owner：Arthur**  
**Architect：GPT / 夥伴**  
**Developer：Claude**  
**Independent QA：Gemini**
