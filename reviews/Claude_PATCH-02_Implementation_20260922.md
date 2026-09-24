# NPSBS_IOMS_V1_0_Claude_PATCH-02_Implementation_20260922

**Author AI**: Claude (Developer)
**Baseline**: `candidates/CANDIDATE-01/` (Arthur confirmed unchanged on GitHub, 2026-09-22)
**Contract**: `contracts/NPSBS_IOMS_V1_0_Production_Contract_V1.1.md` (read in full from GitHub)
**Specification**: `decisions/GPT_PATCH_SPEC_20260922.md` (read in full from GitHub)
**Scope discipline**: Targeted patch only. No wholesale rewrite. No Production API rename.
No 21-column schema reintroduced. No business rule invented that Arthur/GPT did not specify.

---

## 1. Baseline

CANDIDATE-01, 17 `.bas` modules, confirmed by Arthur as unchanged since last delivery
(2026-09-15 conversation). This patch is built directly on top of that confirmed state,
not from any older chat-embedded code snippet.

## 2. Contract Used

`Production_Contract_V1.1.md`, all sections read. Specific sections relied on:
S4-S6 (transaction flow), S7 (Card_Master columns), S11 (Card_Usage_Detail columns),
S14 (frmTreatment control names - unchanged, confirmed txtStart/txtEnd/txtTreatmentType/
txtBookingType), S15 (selectable card types), S16 (Production API name lock).

## 3. Modified Modules (6 of 17)

| Module | Reason |
|---|---|
| `modCard.bas` | P0-02, P0-03, P0-04, P1-01, P1-02 |
| `modCardUsage.bas` | P0-05, P0-06 |
| `modTreatment.bas` | P0-01, P0-02 |
| `modSystemCore.bas` | New `EnsureColumnHeader` helper, required by P0-06/P1-02 to safely append headers to already-existing sheets |
| `modText.bas` | 2 new MsgText keys (`BALANCE_GUARD`, `SETTLEMENT_FAILED`) |
| `modFormBuilder.bas` | 1-line removal: 映遊卡 taken out of `frmCard`'s selectable `cboCardType` list (P1-02) |

## 4. Unmodified Modules (11 of 17)

Byte-for-byte identical to CANDIDATE-01 (verified by diff, not by assumption):
`modAudit.bas`, `modCustomer.bas`, `modDashboard.bas`, `modErrorHandler.bas`,
`modFormLabels.bas`, `modInventory.bas`, `modM5_QA.bas`, `modMainMenu.bas`,
`modProduct.bas`, `modSecurity.bas`, `modValidation.bas`.

## 5. P0 Implementation

### P0-01 — Shortage Validation Before Mutation

**Target**: `modTreatment.CreateTreatmentRecord`

**Before**: shortage resolution was only validated inside `HandleShortage`, which ran
*after* `Card_Master` had already been deducted and `Treatment_Record` already written.

**After**: `CreateTreatmentRecord` is now split into two explicit phases:
- Phase 1 (no mutation): resolve customer, validate treatment/payment, resolve card,
  `CalculateDeduction` (pure calc, no write), determine status, and if `shortage > 0`,
  call `ValidateShortageResolution` and (if resolution = Charge) check
  `shortageAmount > 0` — all before `On Error GoTo Fail` is even set.
- Phase 2 (mutation): only entered once Phase 1 completes without error.

If shortage resolution is missing or invalid, the function now raises **before**
touching any worksheet. Nothing to compensate, nothing partially written.

### P0-02 — Treatment Transaction Atomicity (Pseudo-transaction + Compensation)

**Target**: `modTreatment.CreateTreatmentRecord`

Implemented as: Validate -> Resolve Card -> Calculate Deduction (Balance Guard) ->
Create Treatment_Record -> Commit Deduction -> Create Card_Usage_Detail ->
Handle Shortage -> Audit -> Return.

State flags (`recordCreated`, `cardCommitted`, `usageCreated`) track exactly how far
the transaction got. A single `Fail:` handler (entered only once Phase 2 has started)
compensates in reverse order:
1. If a `Card_Usage_Detail` row was created -> `CompensateUsageDetail` (new Rollback row)
2. If `Card_Master` was committed -> `RestoreCardSnapshot` (restores exact pre-transaction
   `UsedSessions`/`RemainingSessions`, both captured by `CalculateDeduction` before any
   write happened)
3. If `Treatment_Record` was created -> `Status` set to `"Voided"` (never deleted, per
   Contract S12/S27 — the row and its audit metadata remain, only its status changes)
4. Original error is re-raised so the caller (UserForm) still sees the failure.

### P0-03 — Explicit Balance Guard

**Target**: `modCard.DeductSessions` and all callers

`DeductSessions` is now a thin backward-compatible wrapper around two new primitives:
- `CalculateDeduction` (pure, no mutation) — computes `beforeRemaining`, `usedSessions`,
  `shortageSessions`, `afterRemaining` and explicitly raises `BALANCE_GUARD` if
  `usedSessions > beforeRemaining` or `afterRemaining < 0`.
- `CommitDeduction` — re-checks the same guard immediately before writing (defense in
  depth against a stale snapshot), then writes.

No `On Error Resume Next` is used anywhere as a substitute for this validation — the
guard is an explicit `If` check with an explicit `Err.Raise`.

### P0-04 — CardID Format

**Target**: `modCard.GenerateCardID`

Changed date format from `"yymmdd"` (6 digits) to `"yyyymmdd"` (8 digits), and the
existing-ID prefix-match length from `Len(prefix)+6` to `Len(prefix)+8` accordingly.
Example: `C20260922001`. This only affects **newly generated** CardIDs — the function
is never called for a card that already has an explicit CardID (see `CreateCard`), so
none of the 24 existing historical `Card_Master` rows (which use an unrelated
ROC-calendar-style ID from the paper migration, e.g. `C1141217001`) are touched.

### P0-05 — Pending Settlement Atomicity

**Target**: `modCardUsage.SettlePending`

Rewritten to the same Validate -> Snapshot -> Commit -> Create -> Mark -> Audit pattern
as P0-02. `CalculateDeduction` (no mutation) runs first and the `NEW_CARD_SHORT` check
now happens against its pure output rather than a separate ad-hoc balance read. On any
failure after `CommitDeduction`, `Card_Master` is restored via `RestoreCardSnapshot` and
any created usage-detail row is compensated via `CompensateUsageDetail`.
`Pending_Shortage.Status` is left `"Pending"` on failure — it is set to `"Settled"` only
after every step succeeds.

### P0-06 — Card_Usage_Detail Traceability

**Target**: `Card_Usage_Detail` schema and `modCardUsage.CreateUsageDetail`

Added `BeforeSessions` (col 13) and `AfterSessions` (col 14), appended after the
existing 12 columns — their physical order and positions are untouched.
`CreateUsageDetail`'s signature now **requires** `beforeSessions` and `afterSessions`
(not Optional, not silently defaulted to 0) so every call site had to be updated
explicitly rather than compiling with silently-wrong data. Both callers
(`modTreatment.CreateTreatmentRecord`, `modCardUsage.SettlePending`) now pass the exact
`beforeRemaining`/`afterRemaining` values produced by the same `CalculateDeduction` call
used for the deduction itself — never recomputed separately.

Because `modSystemCore.EnsureSheet` only writes headers when a sheet's `A1` is empty
(it will not backfill new header cells on an existing 12-column sheet), a new helper
`modSystemCore.EnsureColumnHeader(ws, colIndex, headerText)` was added and is called
explicitly for columns 13/14 inside `CreateUsageDetail`, and for the five new
`Card_Master` columns (see P1-02) inside `modCard`.

## 6. P1 Implementation

### P1-01 — Oldest Eligible Card First

**Target**: `modCard.FindUsableCard`

The selection condition was `If cid="" Or pdate>bestDate Then` (kept the *latest*
matching card). Changed to `If Not found Or pdate<bestDate Then` (keeps the *earliest*
matching card), with an explicit `found` flag replacing the `cid=""` heuristic for
correctness against a theoretical blank CardID. Eligibility conditions themselves
(CustomerID match, PurchaseDate <= TreatmentDate, RemainingSessions > 0) are unchanged.

### P1-02 — CardType Persistence

**Target**: `modCard.CreateCard`, `Card_Master` schema, `modFormBuilder.BuildCard`

`Card_Master` gains 5 appended columns (11-15): `CardType`, `CreatedAt`, `CreatedBy`,
`ModifiedAt`, `ModifiedBy`. Existing columns 1-10 are untouched. `CreateCard` now writes
`CardType` (previously accepted as a parameter but silently discarded — this was a gap
Claude flagged against itself in the Contract V1.0 review, now closed) plus
`CreatedAt`/`CreatedBy` on insert. `CommitDeduction` and `RestoreCardSnapshot` (both of
which mutate `Card_Master`) now also stamp `ModifiedAt`/`ModifiedBy` on every write,
since they are the only other code paths that touch this sheet.

Historical `Card_Master` rows (24 existing rows, `CardStatus="Observed"`) are **not**
backfilled — their new columns 11-15 stay blank. This is intentional per Contract S7/S28:
their `CardType` cannot be known without paper reconciliation and must not be guessed.

**Also fixed while implementing this item**: `frmCard`'s `cboCardType` dropdown
(built by `modFormBuilder.BuildCard`) previously included `TW_CardYingyou()` as a
selectable option, even though `CardTypeToPrefix` already blocks it with a HOLD error
if selected. Per Contract S15 / Spec P1-02, 映遊卡 must not appear in the selectable
list at all, not merely be blocked at submit time. Removed the one `AddItem` line.
`CardTypeToPrefix`'s HOLD behavior for Yingyou is unchanged (defense in depth, kept in
case a future caller passes it programmatically).

### P1-03 — UserForm Control Contract

No control renames were made. `txtStart`, `txtEnd`, `txtTreatmentType`, `txtBookingType`
on `frmTreatment` are confirmed unchanged and match Contract S14 exactly. Full Control
Manifest delivered separately (see Section 12 / companion file).

## 7. Schema Changes

| Sheet | Change | New physical columns |
|---|---|---|
| `Card_Master` | Appended | 11 CardType, 12 CreatedAt, 13 CreatedBy, 14 ModifiedAt, 15 ModifiedBy |
| `Card_Usage_Detail` | Appended | 13 BeforeSessions, 14 AfterSessions |
| `Treatment_Record` | **No change** | Still 22 columns, per Contract S16 |

Existing columns on both changed sheets are untouched in position and meaning.
Full column-by-column manifest in the companion Schema Manifest file.

## 8. Function Signature Changes

```vb
' CHANGED - new required params, existing 4 callers all updated in this patch
' OLD: CreateUsageDetail(treatmentRecordID, cardID, customerID, usageDate,
'                        sessionsUsed, usageType, [shortageSessions], [notes])
' NEW:
Public Function CreateUsageDetail(ByVal treatmentRecordID As String, ByVal cardID As String, _
        ByVal customerID As String, ByVal usageDate As Date, _
        ByVal beforeSessions As Double, ByVal sessionsUsed As Double, ByVal afterSessions As Double, _
        ByVal usageType As String, Optional ByVal shortageSessions As Double = 0, _
        Optional ByVal notes As String = "") As String

' NEW functions (modCard.bas)
Public Function GetCardBalance(ByVal cardID As String, ByRef usedSessions As Double, ByRef remainingSessions As Double) As Boolean
Public Function CalculateDeduction(ByVal cardID As String, ByVal requested As Double, ByRef beforeUsedTotal As Double, ByRef beforeRemaining As Double, ByRef usedSessions As Double, ByRef shortageSessions As Double, ByRef afterRemaining As Double) As Boolean
Public Sub CommitDeduction(ByVal cardID As String, ByVal beforeRemaining As Double, ByVal usedSessions As Double, ByVal afterRemaining As Double)
Public Sub RestoreCardSnapshot(ByVal cardID As String, ByVal snapshotUsedTotal As Double, ByVal snapshotRemaining As Double, ByVal reason As String)

' NEW function (modCardUsage.bas)
Public Function CompensateUsageDetail(ByVal originalUsageDetailID As String, ByVal reason As String) As String

' NEW sub (modSystemCore.bas)
Public Sub EnsureColumnHeader(ByVal ws As Worksheet, ByVal colIndex As Long, ByVal headerText As String)

' UNCHANGED (signature) - internal behavior changed, callers need no update
Public Function DeductSessions(ByVal cardID As String, ByVal requested As Double, ByRef shortage As Double) As Boolean
Public Function CreateTreatmentRecord(...) As String   ' identical parameter list to CANDIDATE-01
Public Function CreateCard(...) As String               ' identical parameter list to CANDIDATE-01
Public Sub SettlePending(ByVal pendingID As String, ByVal newCardID As String, ByVal sessions As Double)
```

`CreateTreatmentRecord` remains the sole Production Treatment transaction entry point.
No `CreateTreatmentTransaction` exists anywhere in this patch.

## 9. Control Manifest

See companion file `reviews/Claude_PATCH-02_Control_Manifest_20260922.md`.
One control removed from selectable data (not a UI control removal): 映遊卡 taken out
of `cboCardType`'s populated list in `frmCard` (see P1-02).

## 10. Static QA

All performed on the actual PATCH-02 `.bas` files in this delivery:

| Check | Result |
|---|---|
| ASCII-safe source (all 17 files) | PASS |
| Sub/Function begin-end balance (all 17 files) | PASS |
| Duplicate Public procedure names across standard modules | PASS (none) |
| `MsgText` key completeness (31 used, 31 defined) | PASS |
| `TW_*` display-string function completeness | PASS |
| Call-site arity vs. signature (parenthesized calls) | PASS |
| Call-site arity vs. signature (paren-less Sub-style calls: `CalculateDeduction`, `CommitDeduction`, `RestoreCardSnapshot`) | PASS, checked manually line-by-line |
| `CreateUsageDetail` call-site parameter order (3 call sites) | PASS, checked manually — `CompensateUsageDetail`'s reversal call intentionally swaps before/after (documented) |
| Control name vs. handler `Me.<control>` cross-reference (all 8 forms) | PASS (only `frmCard` changed, by removing one `AddItem` line, not a control) |

**Static QA verdict: PASS.** This is static analysis only — see Section 13.

## 11. Known Limitations

- `CardStatus` is referenced by Contract S17 as an eligibility condition for card
  selection, but `modCard.FindUsableCard`'s eligibility check (unchanged by this patch)
  does not filter on it — it only checks `RemainingSessions > 0` and `PurchaseDate`.
  This predates PATCH-02 (present in CANDIDATE-01 too) and is flagged here rather than
  silently left for a future patch to rediscover, since P1-01 touched this exact
  function. Not fixed in PATCH-02 because it was not in the P0/P1 target list and
  fixing it would require deciding what `CardStatus` values should gate eligibility
  (Active only? What about a manually-set "Suspended"?) — a decision, not a bug fix.
- `CompensateUsageDetail` is a no-op (returns empty string) if the original
  `UsageDetailID` cannot be found, or if the original row's `SessionsUsed` was somehow
  <= 0. Both are defensive edge cases that should not occur given
  `CreateUsageDetail`'s own guard, but are noted for completeness.
- `RestoreCardSnapshot` and `CommitDeduction` do not use Excel-level record locking —
  Excel/VBA has no such primitive. Given this is a single-user-at-a-time desktop
  workbook (per Contract's stated environment), this is accepted as out of scope, not
  silently ignored.

## 12. UNVERIFIED Items

- **Compile = UNVERIFIED.** This patch has not been compiled in Excel. Arthur must run
  `Debug -> Compile VBAProject` after importing.
- **Runtime = UNVERIFIED.** No macro in this patch has been executed against the live
  workbook.
- **UAT = UNVERIFIED.** No end-to-end scenario (including the P0-01 shortage-before-save
  scenario and the P0-02 forced-failure compensation scenario) has been run.

None of the above are claimed as PASS anywhere in this report.

## 13. Compile Requirement

Arthur must, in this exact order:
1. Back up the current `.xlsm`.
2. In the VBA editor, open each of the 5 changed modules and replace their content
   with the corresponding file in `candidates/PATCH-02/VBA_Modules/` (select all,
   delete, paste). Do **not** re-import the 12 unmodified modules — they are provided
   only for a complete, self-contained package and for SHA-256 verification.
3. `Debug -> Compile VBAProject`. Report the exact result (PASS, or the first error
   with its module/line).
4. Only after Compile passes: run `BuildProductionUserForms` (rebuilds all 8 forms,
   picking up the one-line `frmCard` change), then `Debug -> Compile VBAProject` again.

## 14. Runtime/UAT Requirement

Suggested minimum scenarios for Arthur/Gemini before this leaves candidate status:
1. Normal treatment, card has enough sessions -> `AutoDeducted`, one `Card_Usage_Detail`
   row with correct Before/Used/After.
2. Treatment where card is short, resolution = Charge, amount > 0 -> succeeds.
3. Same as (2) but resolution left blank -> must fail with the shortage message and
   leave **zero** trace in `Treatment_Record`/`Card_Master`/`Card_Usage_Detail` (this is
   the P0-01 regression test).
4. Force a failure between `CommitDeduction` and `CreateUsageDetail` (e.g., temporarily
   rename `Card_Usage_Detail` sheet) and confirm: `Card_Master` balance is restored,
   `Treatment_Record.Status` = `"Voided"`, `Audit_Log` shows the `COMPENSATE` entries
   (this is the P0-02 regression test — Claude cannot run this without Excel).
5. `SettlePending` happy path and forced-failure path (same idea as 3/4, for P0-05).
6. Purchase a new 彩綿卡/華夏尊盈卡/闔家萬吉卡 card and confirm `CardID` format is
   `PREFIX+YYYYMMDD+###` and `Card_Master` columns 11-15 are populated.
7. Confirm 映遊卡 no longer appears in `frmCard`'s card-type dropdown.
8. Customer with 2 eligible cards of different purchase dates -> confirm the older one
   is selected (P1-01 regression test).

## 15. SHA-256 Manifest

See companion file `candidates/PATCH-02/MANIFEST_SHA256.txt`.

## 16. Discrepancies Between Markdown and Actual Source (Contract S2 compliance)

One discrepancy found and recorded per instructions (not corrected in the Contract —
that is GPT's call):

- This report's own module count in an early internal draft said "5 modified modules."
  While writing Section 4, re-verification against the actual diff showed
  `modSystemCore.bas` is also modified (for `EnsureColumnHeader`), making it 6.
  Corrected in Section 3/4 above rather than left inconsistent. Flagging this
  explicitly because it is exactly the class of error (miscounting/mislabeling
  against actual source) this whole PATCH-02 process exists to prevent.
