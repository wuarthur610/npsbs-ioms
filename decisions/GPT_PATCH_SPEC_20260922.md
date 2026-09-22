# GPT PATCH-02 Specification — NPSBS IOMS V1.0

**Date:** 2026-09-22
**Baseline:** CANDIDATE-01
**Contract:** contracts/NPSBS_IOMS_V1_0_Production_Contract_V1.1.md
**Architect:** GPT
**Developer:** Claude
**QA:** Gemini
**Status:** READY FOR CLAUDE IMPLEMENTATION

## 1. Scope

Implement only the items listed here. Do not perform a wholesale rewrite and do not change locked business rules.

## 2. P0-01 — Validate shortage before mutation

Target: modTreatment.CreateTreatmentRecord and supporting validation.

Required:
- Resolve/validate shortage before DeductSessions.
- If shortage > 0, ValidateShortageResolution must succeed before any Treatment_Record or Card_Master write.
- If resolution is charge, shortageAmount must be > 0 before mutation.
- Normal no-shortage flow must remain unchanged.

Acceptance:
- Invalid shortage resolution produces zero Card_Master mutation.
- Invalid shortage fee produces zero persistent mutation.

## 3. P0-02 — Treatment pseudo-transaction / compensation

Target: modTreatment, modCardUsage, modAudit as required.

Required:
- Capture Card_Master pre-transaction snapshot before first mutation.
- Do not silently delete a Treatment_Record after failure.
- If Treatment_Record was created and a later step fails, mark Status = Voided.
- Restore Card_Master to snapshot state.
- Compensate any newly created usage record with UsageType = Rollback, or use an equivalent audit-safe mechanism explicitly documented in the implementation report.
- Write Audit_Log for failure and compensation.

Do not claim native database transactions exist in Excel/VBA.

## 4. P0-03 — Explicit balance guard

Target: modCard.DeductSessions and all callers.

Required:
- Before mutation calculate BeforeSessions.
- Reject if requested deduction exceeds allowed balance unless the approved shortage flow explicitly supports partial deduction.
- Never permit RemainingSessions < 0.
- Do not rely on On Error Resume Next for business validation.

Acceptance:
- exact balance succeeds;
- partial shortage follows approved resolution;
- negative balance attempt is rejected with no negative write.

## 5. P0-04 — CardID format

Target: modCard.GenerateCardID.

Required format:
`PREFIX + YYYYMMDD + ###`

Examples:
- C20260915001
- H20260915001
- W20260915001

Daily sequence begins at 001.
Do not retain yymmdd generation.

## 6. P0-05 — Pending Settlement atomicity

Target: modCardUsage.SettlePending.

Required:
- Validate target card and balance before persistent settlement writes.
- Snapshot target card.
- Create usage detail and deduct card as one compensatable operation.
- Mark Pending_Shortage Settled only after successful deduction.
- On failure, restore card, compensate usage, and leave Pending_Shortage Pending.

Acceptance:
- no partial settlement state can remain.

## 7. P0-06 — Usage balance traceability

Target: Card_Usage_Detail schema and CreateUsageDetail.

Append fields:
- BeforeSessions
- AfterSessions

Keep existing first 12 physical columns unchanged.

CreateUsageDetail must receive/store:
- BeforeSessions
- SessionsUsed
- AfterSessions

Treatment deduction must write these values from the same snapshot used for the transaction.

## 8. P1-01 — Oldest eligible card

Target: modCard.FindUsableCard.

Current defect: implementation selects the latest eligible PurchaseDate.

Required:
Select the oldest eligible PurchaseDate among:
- matching CustomerID;
- PurchaseDate <= TreatmentDate;
- active CardStatus;
- RemainingSessions > 0.

## 9. P1-02 — CardType persistence

Target: modCard, Card_Master schema, frmCard builder.

Required:
- Persist CardType for new cards.
- Preserve existing first 10 Card_Master columns.
- Append new audit/business fields.
- frmCard must use cboCardType.
- Production selectable card types: 彩綿卡 / 華夏尊盈卡 / 闔家萬吉卡.
- 映遊卡 remains unavailable until Arthur confirms its prefix.

## 10. P1-03 — Control Manifest

Target: modFormBuilder and implementation report.

Verified current CANDIDATE-01 frmTreatment controls include txtTreatmentType and txtBookingType, not cboTreatmentType/cboBookingType. The Gemini finding claiming the opposite is NOT REPRODUCIBLE against the current source.

Do not change these controls solely to satisfy the stale Gemini finding.

Provide a Control Manifest listing every production form and its generated controls.

## 11. Important implementation constraints

- Production transaction entry point remains CreateTreatmentRecord(...).
- Do not reintroduce CreateTreatmentTransaction as a second Production API.
- No real customer names in GitHub artifacts.
- Do not guess 映遊 prefix.
- Do not modify historical package definitions.
- Do not silently overwrite historical records.
- Do not perform a wholesale rewrite.

## 12. Required Claude delivery evidence

Claude must return:
1. Modified module list.
2. Unmodified module list.
3. Exact function signatures.
4. Treatment_Record schema manifest.
5. Card_Master schema manifest.
6. Card_Usage_Detail schema manifest.
7. UserForm Control Manifest.
8. Static QA result.
9. SHA-256 manifest where practical.
10. Explicit note of any remaining UNVERIFIED item.

## 13. Gemini QA priorities after delivery

Gemini must independently verify:
- validation-before-mutation;
- treatment compensation;
- pending settlement compensation;
- balance guard;
- CardID YYYYMMDD format;
- oldest-card selection;
- CardType persistence;
- Before/After usage traceability;
- actual UserForm controls vs handlers.

## 14. Arthur verification gate

After Claude delivery, Arthur should run:

1. Debug → Compile VBAProject
2. BuildProductionUserForms
3. ApplyTraditionalChineseLabels
4. Normal Treatment UAT
5. Shortage UAT
6. Pending Settlement UAT
7. Failure / compensation UAT
8. CardID UAT
9. Oldest-card UAT
10. Balance Guard UAT

## 15. Definition of PATCH-02 complete

PATCH-02 is not complete when Claude finishes coding.

Completion requires:

Claude Implementation → Gemini QA → Arthur Compile/UAT → GPT Final Architecture Gate

**PATCH-02 status: READY FOR IMPLEMENTATION**