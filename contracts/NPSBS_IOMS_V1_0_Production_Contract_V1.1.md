# NPSBS IOMS V1.0 — Production Contract V1.1

**Contract ID:** NPSBS-IOMS-V1.0-PC-V1.1
**Revision Date:** 2026-09-22
**Status:** APPROVED FOR IMPLEMENTATION / PRODUCTION RELEASE BLOCKED
**Architect:** GPT
**Product Owner:** Arthur
**Developer:** Claude
**Independent QA:** Gemini

> This is the normative engineering contract for the next implementation cycle.
> GitHub engineering artifacts must contain no real customer names; use CustomerID or synthetic identifiers.

## 1. Source of Truth

Use this priority when artifacts conflict:
1. Approved Production Contract
2. GPT Architecture Decision
3. Actual GitHub source code
4. Actual XLSM/XLSX binary evidence
5. Claude implementation report
6. Gemini QA report
7. Older handoffs and package descriptions

Actual source evidence overrides descriptive claims.

## 2. Baselines

- FIXED14 = verified technical baseline.
- CANDIDATE-01 = development candidate only; not Production Ready.
- Legacy 21-column MainSystem package = obsolete/archive only.

## 3. Treatment_Record — Locked Target

Treatment_Record is exactly 22 columns:

| # | Field |
|---:|---|
| 1 | TreatmentRecordID |
| 2 | TreatmentDate |
| 3 | TreatmentStart_Taipei |
| 4 | TreatmentEnd_Taipei |
| 5 | CustomerName |
| 6 | CustomerID |
| 7 | Therapist |
| 8 | TherapistID |
| 9 | Sessions |
| 10 | TreatmentType |
| 11 | BookingType |
| 12 | CardID |
| 13 | CardUsageStatus |
| 14 | ShortageResolution |
| 15 | ShortageSessions |
| 16 | TreatmentRevenue |
| 17 | PaymentMethod |
| 18 | CalendarUID |
| 19 | Notes |
| 20 | CreatedAt |
| 21 | CreatedBy |
| 22 | Status |

A VBA function does not need one parameter per database column. Derived/system fields may be generated internally.

Production transaction entry point: CreateTreatmentRecord(...).

## 4. Mandatory Validation Before Mutation

All business validation must finish before the first persistent mutation.

Required validation includes:
- customer / CustomerID resolution
- therapist
- treatment date/time
- sessions
- treatment type
- booking type
- payment method where required
- card eligibility
- shortage resolution when shortage > 0
- shortage amount when charge resolution is selected
- balance guard

CRITICAL RULE: ValidateShortageResolution must execute before any Card_Master deduction or Treatment_Record write.

This closes the source-level defect identified by Claude as FIND-01.

## 5. Treatment Transaction Atomicity

Excel/VBA has no native database transaction across these worksheets. Production therefore requires an explicit pseudo-transaction / compensating transaction.

Required conceptual flow:

START
→ Validate all inputs
→ Resolve Customer / Card
→ Balance Guard
→ Capture Snapshot
→ Create Treatment_Record
→ Create Card_Usage_Detail
→ Update Card_Master
→ Handle approved shortage state
→ Audit
→ COMMIT

No persistent mutation may occur before validation and snapshot.

## 6. Failure / Compensation

If a transaction fails after a persistent write:
- Treatment_Record must not be silently deleted.
- Card_Master must be restored to the pre-transaction state.
- Dependent usage changes must be compensated.
- Audit_Log must record failure and compensation.
- Treatment_Record.Status must become Voided when a transaction row has already been created.

Use compensating transaction / VOID marking, not rollback-by-deletion.

Example:
- Treatment_Record.Status = Voided
- Card_Usage_Detail.UsageType = Rollback
- Card_Master = snapshot state
- Audit_Log = failure + compensation

## 7. Card_Master — Production Requirements

New production cards must support:
- CardID
- CustomerID
- CardType
- PurchaseDate
- PurchaseAmount
- OriginalSessions
- UsedSessions
- RemainingSessions
- PaymentMethod
- CardStatus
- Notes
- CreatedAt
- CreatedBy
- ModifiedAt
- ModifiedBy

New fields may be appended without changing the existing first 10 column positions.

Historical Card_Master rows must not be guessed or retroactively classified. Unknown historical CardType remains unresolved until source evidence is available.

## 8. CardType / CardID Authority

CardType is a business field and must be centrally handled by modCard.

| CardType | Prefix | Status |
|---|---|---|
| 彩綿卡 | C | LOCKED |
| 華夏尊盈卡 | H | LOCKED |
| 闔家萬吉卡 | W | LOCKED |
| 映遊卡 | — | HOLD |

UserForms must not implement independent prefix rules.

Production CardID format:

PREFIX + YYYYMMDD + ###

Examples:
- C20260915001
- H20260915001
- W20260915001

Daily sequence starts at 001.
The current yymmdd format is not compliant and must be corrected.

## 9. Card Selection

Eligible card:
- CustomerID matches
- PurchaseDate <= TreatmentDate
- CardStatus active
- RemainingSessions > 0

When multiple cards are eligible, consume the oldest eligible card first.
A later-purchased card must not be consumed while an older eligible card still has remaining sessions.
No retrospective deduction against a later-purchased card is allowed.

## 10. Balance Guard

Before deduction:

BeforeSessions >= SessionsToDeduct

After deduction:

AfterSessions >= 0

Any deduction that would create a negative balance must be rejected before Card_Master is written.
On Error Resume Next is not a business-rule safeguard.

## 11. Card_Usage_Detail

Production traceability must answer:
- Which treatment caused this deduction?
- What was the balance before and after?

Minimum logical fields:
1. UsageDetailID
2. TreatmentRecordID
3. CardID
4. CustomerID
5. UsageDate
6. BeforeSessions
7. SessionsUsed
8. AfterSessions
9. UsageType
10. ShortageSessions
11. Notes
12. CreatedAt
13. CreatedBy
14. Version

BeforeSessions and AfterSessions should be appended so existing first-12-column indexes are not disturbed.

## 12. Shortage Rules

Approved shortage resolutions:
1. Charge for reduced sessions
2. Direct waiver
3. Pending settlement on a future card purchase

Rules:
- no negative balance;
- no silent treatment rewrite;
- no retrospective deduction;
- original Treatment_Record remains historically identifiable;
- pending settlement creates a new usage transaction.

## 13. Pending Settlement Atomicity

Required flow:

Validate → Snapshot → Create settlement UsageDetail → Deduct Card_Master → Mark Pending_Shortage settled → Audit → Commit

If any step fails:
- restore Card_Master;
- compensate the newly created settlement transaction;
- do not mark Pending_Shortage as settled.

Partial settlement is prohibited.

## 14. UserForm Control Contract

Control names are part of the implementation contract.

Verified frmTreatment names include:
- txtStart
- txtEnd
- txtTreatmentType
- txtBookingType

Obsolete aliases txtStartTime and txtEndTime must not be used unless the actual Builder is changed and the Control Manifest is updated.

Gemini QA must compare:
1. controls created by modFormBuilder;
2. controls referenced by event handlers;
3. controls referenced by SaveTreatmentForm;
4. expected data types.

## 15. frmCard Control Contract

frmCard must expose cboCardType.

Production selectable values:
- 彩綿卡
- 華夏尊盈卡
- 闔家萬吉卡

映遊卡 remains HOLD until its prefix is formally confirmed.

The form must not hard-code C as the prefix. CardID generation must call the central Card module.

## 16. Package Master — Deferred

A future Package_Master should version:
- CardType
- price
- product entitlement
- original sessions
- discount
- EffectiveFrom
- EffectiveTo
- Status

This is deferred to PATCH-03 or later unless Arthur promotes it.
Historical package definitions must never be overwritten by new definitions.

## 17. Audit / Versioning

Material corrections must preserve:
- Who
- When
- What
- Old Value
- New Value
- Reason

Historical information must not be silently erased.

## 18. Calendar Authority

Google Calendar remains authoritative for actual appointment/treatment timing.

Rules:
- Appointment is not Treatment.
- Cancellation retains history.
- Reschedule creates a new event while retaining the cancelled event.
- All-day closure/leave is not treatment.
- Unconfirmed Calendar-vs-Excel discrepancies remain Review/Staging.
- Revenue must not be inferred from an appointment.

## 19. Historical Migration

Historical migration scope: 2025/12 through 2026/08.

Every record is classified as Confirmed, Review, or Unmatched.
Review and Unmatched records must not be silently promoted to Production.

## 20. QA Evidence Standard

Every major finding must include:
- Severity
- Artifact
- Version
- Module
- Procedure
- Source location
- Observed behavior
- Expected behavior
- Reproduction status
- Recommendation

Allowed status labels:
CRITICAL / HIGH / MEDIUM / LOW / CONFLICT / UNVERIFIED / REJECTED-NOT-REPRODUCIBLE

Hypothesis without runtime evidence must not be represented as confirmed runtime fact.

## 21. Version Evidence

Every Candidate release must provide:
- Candidate name
- Date
- Baseline
- Modified modules
- Unmodified modules
- Schema version
- Build status
- Compile status
- Runtime status
- SHA-256 manifest where practical

Recommended artifact: MANIFEST_SHA256.txt

## 22. Definition of Done

Designed → Implemented → Static QA Passed → Compile Passed → Runtime Passed → Gemini QA Passed → GPT Gate Passed → Arthur UAT Passed → Production Candidate → Production

Claude implementation alone does not mean Done.

## 23. AI Governance

- Arthur = Product Owner / final business UAT authority.
- GPT = System Architect / contract authority / architecture gate.
- Claude = Developer / implementation authority within the contract.
- Gemini = Independent QA / adversarial source reviewer.

No AI may unilaterally change locked business rules.

## 24. No Wholesale Rewrite

Claude must not perform a wholesale rewrite in response to an isolated QA finding.

Required workflow:
GPT Contract → GPT Patch Specification → Claude Patch → Gemini QA → Arthur Compile/UAT → GPT Gate

## 25. PATCH-02 Priority

### P0
1. Move shortage validation before first mutation.
2. Implement transaction snapshot/compensation.
3. Implement explicit balance guard.
4. Correct CardID format.
5. Prevent partial pending settlement.
6. Add Card_Usage_Detail BeforeSessions / AfterSessions.

### P1
7. Correct oldest-eligible-card selection.
8. Persist CardType for new production cards.
9. Align Control Manifest with actual Builder.

### P2 / Deferred
10. Package_Master and package-effective-date framework.

## 26. Production Release Gate

Production is BLOCKED until all are PASS:
- Contract V1.1 implemented
- P0 items implemented
- Static source checks pass
- Excel Debug → Compile VBAProject = PASS
- BuildProductionUserForms = PASS
- ApplyTraditionalChineseLabels = PASS
- Normal Treatment UAT = PASS
- Shortage UAT = PASS
- Pending Settlement UAT = PASS
- Failure / compensation UAT = PASS
- CardID UAT = PASS
- Oldest-card UAT = PASS
- Balance Guard UAT = PASS
- Card_Usage_Detail traceability UAT = PASS
- Gemini QA = PASS
- GPT Architecture Gate = PASS
- Arthur final UAT = PASS

## 27. GitHub Privacy Rule

No real customer names may appear in contracts, decisions, reviews, test cases, migration reports, engineering screenshots, or sample data.
Use CustomerID or synthetic identifiers.

## 28. Final Architecture Principle

> Business rules belong in the system core, not scattered across UserForms.

> Historical facts must be preserved.

> A transaction is successful only when all required persistent changes succeed.

> Failed transactions are compensated and audited; historical records are not silently erased.

> Actual source evidence takes precedence over AI descriptions.

> No AI handoff overrides the approved Production Contract.

## 29. Required Next Actions

### Claude
- Acknowledge V1.1.
- Implement PATCH-02 only against this contract.
- Provide modified/unmodified module manifest.
- Provide schema manifest.
- Provide Control Manifest.
- Provide function signatures.
- Provide SHA-256 manifest.
- Provide implementation report.

### Gemini
- Audit PATCH-02 against V1.1.
- Prioritize shortage validation, transaction compensation, balance guard, CardID, card selection, usage detail, and UserForm controls.

### Arthur
After PATCH-02 delivery, run Excel Compile and the required UAT suite.

### GPT
After Claude + Gemini + Arthur evidence, issue the final Architecture Gate.

**Contract Status: APPROVED FOR IMPLEMENTATION**
**Production Status: BLOCKED**
**Next Milestone: PATCH-02**