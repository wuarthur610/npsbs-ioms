# NPSBS_IOMS_V1_0_Claude_PATCH-02_Schema_Manifest_20260922

Full physical column layout for every sheet touched (directly or indirectly) by
PATCH-02. Sheets not listed here are unaffected by this patch.

## Card_Master — 10 unchanged + 5 appended (15 total)

| # | Column | Status |
|---|---|---|
| 1 | CardID | unchanged |
| 2 | CustomerID | unchanged |
| 3 | PurchaseDate | unchanged |
| 4 | PurchaseAmount | unchanged |
| 5 | OriginalSessions | unchanged |
| 6 | UsedSessions | unchanged (now also updated by `RestoreCardSnapshot`) |
| 7 | RemainingSessions | unchanged (now also updated by `RestoreCardSnapshot`) |
| 8 | PaymentMethod | unchanged |
| 9 | CardStatus | unchanged |
| 10 | Notes | unchanged |
| 11 | **CardType** | **NEW (PATCH-02)** — written by `CreateCard`; blank on the 24 pre-existing historical rows (not backfilled) |
| 12 | **CreatedAt** | **NEW (PATCH-02)** — written by `CreateCard` only |
| 13 | **CreatedBy** | **NEW (PATCH-02)** — written by `CreateCard` only |
| 14 | **ModifiedAt** | **NEW (PATCH-02)** — written by `CreateCard`, `CommitDeduction`, `RestoreCardSnapshot` |
| 15 | **ModifiedBy** | **NEW (PATCH-02)** — written by `CreateCard`, `CommitDeduction`, `RestoreCardSnapshot` |

Header cells for columns 11-15 are written by `modCard.EnsureCardMasterAuditColumns`
(via `modSystemCore.EnsureColumnHeader`) the first time any of `CreateCard`,
`CommitDeduction`, or `RestoreCardSnapshot` runs after this patch is applied — they do
not require a separate migration step.

## Card_Usage_Detail — 12 unchanged + 2 appended (14 total)

| # | Column | Status |
|---|---|---|
| 1 | UsageDetailID | unchanged |
| 2 | TreatmentRecordID | unchanged |
| 3 | CardID | unchanged |
| 4 | CustomerID | unchanged |
| 5 | UsageDate | unchanged |
| 6 | SessionsUsed | unchanged |
| 7 | UsageType | unchanged — now also takes the value `"Rollback"` (new, written by `CompensateUsageDetail`) in addition to the existing `"TreatmentDeduction"` / `"PendingSettlement"` |
| 8 | ShortageSessions | unchanged |
| 9 | Notes | unchanged |
| 10 | CreatedAt | unchanged |
| 11 | CreatedBy | unchanged |
| 12 | Version | unchanged (still always written as `1`) |
| 13 | **BeforeSessions** | **NEW (PATCH-02)** |
| 14 | **AfterSessions** | **NEW (PATCH-02)** |

Header cells for columns 13-14 are written by `modCardUsage.CreateUsageDetail` itself
(via `EnsureColumnHeader`) on first use after this patch is applied.

## Treatment_Record — 22 columns, NO CHANGE

Per Contract V1.1 S16, this patch does not alter `Treatment_Record`'s schema. Column 22
(`Status`) now takes the value `"Voided"` in addition to the existing `"Completed"` —
this is a new **value**, not a new **column**.

## Pending_Shortage — 11 columns, NO CHANGE

Schema unchanged. `Status` continues to use `"Pending"` / `"Settled"` only; PATCH-02's
compensation logic for `SettlePending` never marks a pending row `"Settled"` unless the
full sequence (deduct -> usage detail -> mark) succeeds, so no new status value was
needed here.

## Audit_Log — 10 columns, NO CHANGE

Schema unchanged. `ActionType` now also takes the value `"COMPENSATE"` in addition to
the existing `"CREATE"` / `"UPDATE"` — again a new value, not a new column.
