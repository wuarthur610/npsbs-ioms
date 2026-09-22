# 目前狀態

## 2026-09-22 — GPT Production Contract V1.1 Locked

**Status:** APPROVED FOR IMPLEMENTATION / PRODUCTION BLOCKED

- FIXED14 = verified technical baseline.
- CANDIDATE-01 = development candidate; not Production Ready.
- GPT Decision: decisions/GPT_Decision_20260915.md
- Production Contract V1.1: contracts/NPSBS_IOMS_V1_0_Production_Contract_V1.1.md
- V1.1 is now the normative implementation contract.
- Claude may proceed with PATCH-02 only against V1.1.
- Gemini must independently QA the resulting PATCH-02 candidate.
- Arthur remains final Compile/UAT authority.

### V1.1 P0 blockers
1. Shortage validation before any persistent mutation.
2. Treatment transaction snapshot + compensating transaction / VOID marking.
3. Explicit card balance guard.
4. CardID format PREFIX + YYYYMMDD + ###.
5. Pending Settlement atomicity.
6. Card_Usage_Detail BeforeSessions / SessionsUsed / AfterSessions.

### V1.1 P1
1. Oldest eligible card first.
2. CardType persistence for new production cards.
3. UserForm Control Manifest alignment.
4. 映遊卡 remains HOLD until prefix is confirmed.

### Deferred
- Package_Master / effective-date package versioning = PATCH-03 or later unless Arthur promotes it.

### Resolved findings
- Gemini CRITICAL-01: 22 DB columns vs 16 function parameters = REJECTED as a contract violation.
- Gemini control mismatch finding based on older artifact = REJECTED / NOT REPRODUCIBLE against current evidence.
- Claude FIND-01 = ACCEPTED and promoted to P0.
- Claude CONFLICT-01 = ACCEPTED; add CardType/audit fields without disturbing existing first 10 positions.
- Claude CONFLICT-02 = ACCEPTED; append BeforeSessions/AfterSessions.
- Claude CONFLICT-03 = resolved by V1.1 transaction contract.
- Claude CONFLICT-04 = resolved by compensating transaction / VOID marking, not deletion.
- Claude FIND-02 = accepted; use verified txtStart / txtEnd.
- Claude FIND-03 Package_Master = deferred.

### Next workflow
GPT Contract V1.1 → Claude PATCH-02 → Gemini QA → Arthur Compile/UAT → GPT Final Gate → Production

### Current action owners
- [ ] Claude: acknowledge V1.1 and implement PATCH-02.
- [ ] Gemini: perform independent PATCH-02 QA.
- [ ] Arthur: after PATCH-02 delivery, run Excel Compile + UAT.
- [ ] GPT: final architecture gate after evidence.

### Privacy
No real customer names in GitHub engineering artifacts. Use CustomerID or synthetic IDs.

## 2026-09-15 — GPT Architecture Decision

**Status:** CONDITIONAL APPROVAL

- FIXED14 remains the verified technical baseline.
- CANDIDATE-01 remains a development candidate and is NOT Production Ready.
- GPT Decision: decisions/GPT_Decision_20260915.md
- Production Contract V1.1 supersedes the V1.0 draft for implementation.
- Production remains blocked pending P0 fixes and verification.

## 目前有效版本

| 項目 | 檔案 | 狀態 |
|---|---|---|
| 已驗證基線 | FIXED14 | Compile PASS / Forms PASS |
| 最新候選 | candidates/CANDIDATE-01/ | Candidate / not Compile-verified |
| GPT Decision | decisions/GPT_Decision_20260915.md | Conditional Approval |
| Latest Contract | contracts/NPSBS_IOMS_V1_0_Production_Contract_V1.1.md | APPROVED FOR IMPLEMENTATION |
| Previous Contract | contracts/NPSBS_IOMS_V1_0_Production_Contract_V1.0.md | Historical / superseded |

## 已作廢，請勿使用
- archive/MainSystem_Package_OBSOLETE/ — 21-column legacy package.

## 待 Arthur 決定
1. 映遊卡正式 CardID prefix。
2. 套卡新規格正式生效日；歷史卡不得回溯套用新規格。