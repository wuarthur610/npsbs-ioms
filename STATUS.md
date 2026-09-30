# 目前狀態

## 2026-09-29 — PATCH-02 Hotfix-02 Complete (Claude)

- 觸發：Arthur TEST-05 實機測試，14 個欄位中 5 個 FAIL（調理師下拉、堂數自動計算、
  治療類型錯字+無下拉、預約類型無下拉、套卡號無內容）
- Root Cause（對照實際原始碼確認）：cboTherapist/cboCardID 建了控制項但從未 AddItem；
  txtStart/txtEnd 完全沒有事件處理常式；txtTreatmentType/txtBookingType 是純
  TextBox 無法有下拉；modFormLabels.bas 有一個字元編碼打錯（療療類型應為治療類型）
- 決策依據：GPT 2026-09-29 正式裁決 3 項開放問題（Control Contract 只鎖 Name 不鎖
  控制項類型／治療類型與預約類型選項清單／Therapist_Master 欄位結構），Claude 完全
  依裁決執行，未自行假設
- 修改模組（4／17）：modFormBuilder.bas、modCard.bas、modSystemCore.bas、
  modFormLabels.bas；其餘 13 個逐位元組未動，modTreatment.bas（交易核心）完全未碰
- 新增邏輯：LoadTherapistList（依 Header 讀 Therapist_Master，不寫死欄位位置）、
  RefreshCardList（CustomerID 變動後重新載入套卡下拉，重用既有 FindUsableCard
  資格邏輯，未建立第二套）、RecalcSessions（重用既有 CalculateSessions，未建立
  第二套堂數計算）
- 額外完成：SaveTreatmentForm 的 TherapistID 參數原本寫死空字串，這次一併接上
  （直接對應「保留 Therapist/TherapistID 現有架構」要求，非擴大範圍，Report 已載明）
- Static QA：PASS（全 17 檔 ASCII／Sub-End 配對／無重複宣告／跨模組呼叫全部有
  定義／換行接續符號未超標）
- Compile／Runtime／UAT：UNVERIFIED
- Artifacts:
  * `candidates/PATCH-02/VBA_Modules/`（4 個檔案待覆蓋：modFormBuilder.bas /
    modCard.bas / modSystemCore.bas / modFormLabels.bas）
  * `reviews/Claude_PATCH-02_Hotfix-02_Implementation_20260929.md`
  * `MANIFEST_SHA256.txt`（全 17 檔雜湊，供核對實際改了哪幾個）

**最後更新**：2026-09-29（Arthur）

## 等誰動作

- [ ] Gemini：獨立 QA（任務文件第 9 節有明確檢查清單，特別是 DataSource 與
  CardID 篩選正確性）
- [ ] GPT：Gemini QA 後做 Architecture Gate
- [ ] Arthur：覆蓋 4 個模組 → Compile → BuildProductionUserForms → 重跑 TEST-05
- [ ] Claude：暫停寫程式，等 Gemini QA + Arthur TEST-05 結果

## 待 Arthur 決定的事

1. 映遊卡的 CardID 前綴是什麼？
2. 套卡新規格從哪一天開始生效？（舊卡不得回溯套用）
3. Therapist_Master 的 Status 欄位實際存的是英文「Active」還是中文「在職」？
   （這次寫成兩種都接受，TEST-05 可以順便驗證實際是哪一種）
