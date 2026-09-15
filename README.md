# NPSBS IOMS V1.0

如印植秘三峽北大店 營運管理系統。本 repo 是 Arthur、GPT、Claude、Gemini 的共同工作區。

## 角色
- Arthur：Product Owner / 最終 UAT
- GPT：System Architect / Technical Gatekeeper
- Claude：Developer / Implementer
- Gemini：Independent QA / Code Reviewer

## 資料夾
| 資料夾 | 內容 | 誰放的 |
|---|---|---|
| `contracts/` | Production Contract（最高規範） | GPT |
| `candidates/` | VBA 原始碼候選版本 | Claude |
| `reviews/` | 審查報告 | Claude / Gemini |
| `decisions/` | 架構裁決 | GPT |
| `archive/` | 已作廢版本（保留不刪） | 任何人 |

## 鐵則
1. **絕不上傳 `.xlsm`**，也不上傳任何含真實客戶資料的檔案。
2. 文件中一律使用 CustomerID，**不得出現客戶姓名**。
3. 任何 AI 發言前，必須指出所依據的**檔案路徑**。
4. 實際檔案 > 文件描述。文件寫的不算數，repo 裡的檔案才算數。
5. `archive/` 的東西永遠不刪，只標記作廢。

## 目前狀態
見 `STATUS.md`
