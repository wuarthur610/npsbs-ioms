# NPSBS_IOMS_V1_0_Claude_PATCH-02_CrossModuleCheck_20260925

**觸發原因**：Arthur 在 TEST-02 Compile 時，`modCard.bas` 的
`EnsureColumnHeader ws, 11, "CardType"` 一行報「Sub 或 Function 未定義」。

**結論**：不是原始碼包遺漏這個函式。是 Arthur 現有的 Excel 工作簿裡，
`modSystemCore` 模組還停留在沒有 `EnsureColumnHeader` 的舊版本，而 `modCard`
已經是會呼叫它的新版本，兩邊版本沒對齊。

## 根本原因

Claude 在 2026-09-22 的 PATCH-02 Implementation Report 第 3/4 節裡，先寫
「5 個模組修改」，隨後才在文字中自我更正為「其實是 6 個，漏算了
`modSystemCore.bas`」。這個更正沒有用足夠醒目的方式呈現，很可能導致依清單
匯入時，`modSystemCore.bas` 被跳過。這是 Claude 報告呈現方式的責任，
不是 Arthur 操作疏失。

## 驗證過程（依 Arthur 上傳的 `Claude_20260925_patch02_VBA_Modules.zip`，17 個檔案）

1. **`EnsureColumnHeader` 是否存在** —— 存在，`modSystemCore.bas` 第 34 行
   `Public Sub EnsureColumnHeader(ByVal ws As Worksheet, ByVal colIndex As Long, ByVal headerText As String)`。
   呼叫端：`modCard.bas`（5 處）、`modCardUsage.bas`（2 處），呼叫方式一致。

2. **應歸屬模組** —— `modSystemCore.bas`（原本就存放 `EnsureSheet`／`LastDataRow`／
   `NextID`／`CurrentUser`／`SafeNumber` 等共用工作表工具函式的模組）。不需要搬遷。

3. **是否需要補回** —— 不需要，這份包裡沒有缺漏。

4. **全庫「呼叫但未定義」交叉檢查**（17 個檔案，兩輪）：
   - 括號式呼叫（`Identifier(...)`）：排除 VBA/Excel 內建關鍵字與函式後，
     零筆未定義。
   - 不帶括號的 Sub 式呼叫（如 `CommitDeduction actualCard, beforeRemain, ...`）：
     先合併換行接續符號再逐句解析，零筆未定義。

5. **Static QA 重跑結果**：

   | 檢查項目 | 結果 |
   |---|---|
   | ASCII-safe（全部 17 檔） | PASS |
   | 跨檔案重複 Sub/Function 宣告 | PASS（無重複） |
   | 每檔 Sub/End、Function/End 配對 | PASS（全部平衡） |
   | Cross-module reference（呼叫但未定義） | PASS（零筆，見上） |

6. **修正版 `.bas`** —— 本次不需要修正原始碼。附上的
   `modSystemCore.bas` 就是 Arthur 上傳這份包裡驗證通過、原封不動的版本，
   供直接覆蓋 Excel 工作簿裡的 `modSystemCore` 模組使用。

## 給 Arthur 的下一步

不需要重建 `.xlsm`：

1. 開 VBA 編輯器，找到 `modSystemCore` 模組。
2. 全選現有內容 → 刪除。
3. 貼上本次附上的 `modSystemCore.bas` 全部內容。
4. `Debug -> Compile VBAProject`。
5. 若這次 Compile 通過，回到原本的流程：TEST-03 `BuildProductionUserForms`。
6. 若還有其他錯誤，把完整錯誤訊息（模組名 + 行號 + 錯誤文字）貼給 Claude。

## 給 GPT／Gemini 的說明

這不是一個新的程式碼缺陷，也不需要新的 Hotfix 編號——是一次「Claude 交付的模組
清單在文件呈現上不夠清楚，導致部署時漏掉一個檔案」的操作對齊問題。已透過
本次交叉比對確認：現有 PATCH-02 原始碼包（17 檔）內部一致，沒有其他遺漏。
