# NPSBS IOMS V1.0 — PATCH-02 Hotfix-02 修正任務

**Task ID:** PATCH-02-HOTFIX-02  
**Status:** READY FOR CLAUDE IMPLEMENTATION  
**Baseline:** PATCH-02 Hotfix-01  
**Developer:** Claude  
**Independent QA:** Gemini  
**Architecture Gate:** GPT  
**Final UAT:** Arthur

## 1. TEST-05 實機結果

| 欄位 | 結果 |
|---|---|
| 客戶姓名 | PASS |
| CustomerID | PASS；輸入客戶姓名後自動帶出 `TMP-CUST-0047` |
| 調理日期 | REVIEW；預設今日日期，可手動修改 |
| 調理師 | FAIL；有下拉按鈕但無選項 |
| 開始時間 | PASS；可手動輸入 |
| 結束時間 | PASS；可手動輸入 |
| 堂數 | FAIL；未依開始/結束時間自動計算 |
| 治療類型 | FAIL；顯示「療療類型」且無下拉選項 |
| 預約類型 | FAIL；無下拉選項 |
| 套卡號 | FAIL；有下拉按鈕但無內容 |
| 付款方式 | PASS；下拉及兩個選項正常 |
| 療程收入 | PASS；可輸入數字 |
| 不足堂數處理 | PASS；下拉及選項正常 |
| 不足收費 | PASS；可輸入數字 |
| 備註 | PASS；可輸入文字 |

## 2. Hotfix-02 修正項目

### P0-01 調理師下拉
- 依既有 `Therapist_Master` / Master Data 初始化。
- 不得在 `frmTreatment` 硬編碼調理師姓名。
- 保留 Therapist / TherapistID 現有架構。
- 目前既定調理師 Master 為 A–G，新增調理師後應可透過 Master Data 使用。

### P0-02 堂數自動計算
- 開始時間與結束時間輸入後，堂數必須自動計算。
- 既定服務規則：40 分鐘 = 1 堂。
- 10:00–10:40 = 1 堂；10:00–11:20 = 2 堂；10:00–12:00 = 3 堂。
- 優先重用現有 Session Calculation / `CalculateSessions` 邏輯。
- 不建立第二套堂數計算邏輯。
- `txtStart` 或 `txtEnd` 改變後應重新計算。
- 空白、格式錯誤、結束早於開始等情況不得產生錯誤堂數。
- 不得破壞既有 Transaction Core。

### P0-03 治療類型
- 「療療類型」修正為「治療類型」。
- 治療類型必須有可選下拉內容。
- 使用現有 Lists / Master Data / 共用 DataSource。
- 保持 Control Contract：`txtTreatmentType`。
- 不得改成 `cboTreatmentType`。

### P0-04 預約類型
- 初始化既有 Booking Type DataSource。
- 保持 Control Contract：`txtBookingType`。
- 不得改成 `cboBookingType`。
- 不建立第二套 Booking Type 定義。

### P0-05 套卡號
- 依目前 `CustomerID` 查詢 `Card_Master`。
- 僅列出該 CustomerID 的可用卡。
- 基本資格：`CustomerID = 當前 CustomerID`、`PurchaseDate <= TreatmentDate`、`RemainingSessions > 0`。
- 遵守 oldest eligible card first。
- 若已有 `FindUsableCard`，優先重用。
- 不得在 UserForm 建立第二套 Card eligibility 邏輯。
- 不得顯示其他 CustomerID 的卡或 RemainingSessions = 0 的卡。
- CustomerID 改變後，CardID 下拉必須重新載入。
- `CardStatus = Active` 若目前 source 已具備判斷，應使用；若尚未完成，於 Report 標示，不得自行擴大範圍。

## 3. 不得破壞的既有功能
- CustomerName → CustomerID 自動帶入
- 付款方式下拉
- 不足堂數處理下拉
- 療程收入輸入
- 不足收費輸入
- 備註輸入

## 4. 本次不處理
調理日期目前可正常預設今日日期並可手動修改。若 Claude 發現明確 Bug，僅於 Report 提出 Observation；未經 GPT 核准不得擴大範圍。

## 5. 不可修改事項
1. Wholesale Rewrite。
2. Treatment_Record 22 欄 Schema。
3. PATCH-02 Transaction Core。
4. Shortage / Pending Settlement architecture。
5. CardID 格式。
6. Card Usage Before/Used/After 架構。
7. Compensation / Rollback architecture。
8. Production Contract V1.1。
9. `txtTreatmentType` / `txtBookingType` Control Contract。
10. 在 Form 中硬編碼調理師資料。
11. 在 Form 中建立另一套 Card eligibility 邏輯。

## 6. Claude 必須檢查
`modFormBuilder.bas`, `modTreatment.bas`, `modIOMS_Core.bas`, `modCard.bas`, `modValidation.bas`, `modSystemCore.bas`, `modText.bas`，以及 `Therapist_Master`, `Card_Master`, Lists / Master Data。

確認 Control 建立、初始化、DataSource、CustomerID lookup、Card lookup、Session calculation、Treatment Type、Booking Type。

## 7. Static QA
- ASCII-only source = PASS
- Sub / End Sub、Function / End Function、Property / End Property 平衡
- 無 duplicate procedure
- Control Manifest 保持 `txtStart`, `txtEnd`, `txtTreatmentType`, `txtBookingType`
- Report 必須列出 Therapist / Treatment Type / Booking Type / CardID 實際 DataSource
- 不得只寫「已修正」

## 8. Claude 交付
- 所有實際修改的 `.bas`
- `NPSBS_IOMS_V1_0_PATCH-02_Hotfix-02_Package.zip`
- `reviews/Claude_PATCH-02_Hotfix-02_Implementation_202609XX.md`

Report 必須包含：原始問題、Root Cause、修改模組、修改程序、DataSource、Control Manifest、Session Calculation、Card Filter、Static QA、未驗證項目、SHA-256、Git commit。

Claude 不得宣稱未實際執行的 Excel Compile / Runtime / UAT 為 PASS。

## 9. Gemini QA
特別檢查：
- 調理師 DataSource
- 治療類型 DataSource
- 預約類型 DataSource
- CardID 是否依 CustomerID 正確過濾
- 是否錯誤顯示其他客戶卡
- 是否顯示 RemainingSessions = 0 卡
- Session 是否自動計算
- `txtTreatmentType` / `txtBookingType` 是否保持不變
- 是否破壞 PATCH-02 Transaction Core
- 是否新增 duplicate event handler

## 10. Arthur Excel UAT
重新執行 TEST-05：
- 客戶姓名 → CustomerID 自動帶入
- 調理師下拉有有效調理師
- 10:00–11:20 → 堂數 = 2
- 治療類型下拉有選項且顯示「治療類型」
- 預約類型下拉有選項
- CustomerID 對應套卡號下拉有正確可用卡
- 付款方式正常
- 不足堂數處理正常

## 11. Definition of Done
- [ ] 調理師下拉正常
- [ ] 堂數依開始/結束時間自動計算
- [ ] 治療類型下拉正常
- [ ] 「療療類型」修正為「治療類型」
- [ ] 預約類型下拉正常
- [ ] 套卡號依 CustomerID 正確載入
- [ ] CustomerID 自動帶入未被破壞
- [ ] 既有 PASS 功能未被破壞
- [ ] Treatment_Record 22 欄未被破壞
- [ ] PATCH-02 Transaction Core 未被改寫
- [ ] Static QA PASS
- [ ] Gemini QA PASS
- [ ] GPT Gate PASS
- [ ] Excel Compile PASS
- [ ] BuildProductionUserForms PASS
- [ ] Arthur TEST-05 PASS

## 12. Workflow
`Claude Implementation → Gemini Independent QA → GPT Architecture Gate → Arthur Excel UAT → GPT Final Gate`

## 13. Engineering Principle
> PATCH-02 Hotfix-02 是 **UserForm DataSource / Calculation Fix**，不是 Transaction Architecture Rewrite。

Claude 必須採：`Root Cause → Minimum Fix → Static QA → Gemini QA → GPT Gate → Arthur UAT`

未經 GPT 核准，不得擴大 Hotfix-02 範圍。
