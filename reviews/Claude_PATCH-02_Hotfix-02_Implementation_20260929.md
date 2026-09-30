# NPSBS_IOMS_V1_0_Claude_PATCH-02_Hotfix-02_Implementation_20260929

**Author AI**: Claude (Developer)
**Baseline**: PATCH-02 Hotfix-01 (Arthur-confirmed 17-file package, `modSystemCore.bas`
already aligned via the 2026-09-25 cross-module check)
**Decision basis**: GPT's 2026-09-29 formal ruling on the 3 open questions Claude raised
before implementation (control-type vs Name, TreatmentType/BookingType option lists,
Therapist_Master column layout) — reproduced in full in Section 2.
**Scope discipline**: `frmTreatment` DataSource/calculation wiring only. No transaction
core, schema, CardID format, or shortage/pending architecture touched.

---

## 1. Original Problem (TEST-05)

5 of 14 fields failed: 調理師下拉無選項、堂數未自動計算、「療療類型」錯字且無下拉、
預約類型無下拉、套卡號下拉無內容.

## 2. Root Cause (confirmed against actual `modFormBuilder.bas`, not assumed)

| Field | Root cause |
|---|---|
| 調理師 (cboTherapist) | Control built (`AddCombo`) but never populated — zero `AddItem` calls existed anywhere for it |
| 堂數 (txtSessions) | `txtStart`/`txtEnd` had no event handlers at all, so nothing ever triggered a calculation. The calculation function itself (`modTreatment.CalculateSessions`, 40-minute rule) already existed and was correct — it was simply never called from the form |
| 「療療類型」 | `modFormLabels.bas`, the `lblType` caption's `UText` call had the first character code duplicated (療,療,類,型 instead of 治,療,類,型) |
| 治療類型 (txtTreatmentType) | Built as a plain `TextBox` (`AddText`), which cannot have a dropdown list by definition |
| 預約類型 (txtBookingType) | Same as above |
| 套卡號 (cboCardID) | Control built but never populated, same pattern as 調理師 |

No `modIOMS_Core.bas` exists in the current 17-module PATCH-02 set — confirmed by
directory listing. `CalculateSessions` is in `modTreatment.bas`. GPT confirmed this in
its 2026-09-29 ruling as a stale reference to an earlier package naming, not a Claude
error.

## 3. GPT's Formal Ruling (2026-09-29) — followed exactly

1. **Control Contract governs `Name`, not the underlying control type.**
   `txtTreatmentType`/`txtBookingType` keep their Name; their underlying type is changed
   to ComboBox. Renaming to `cboTreatmentType`/`cboBookingType` remains forbidden.
2. **DataSource for 治療類型**: 一般調理, 體驗調理.
   **DataSource for 預約類型**: 預約, Walk-in, 改期, 店長安排.
   No other values, no early-draft values, used.
3. **Therapist_Master columns**: `TherapistID`, `TherapistName`, `Status`, `Notes`.
   Read by header text (not a fixed column position), load only Active/在職 rows,
   preserve the existing Therapist+TherapistID two-field architecture.

## 4. Modified Modules (4 of 17)

| Module | Change |
|---|---|
| `modFormBuilder.bas` | `BuildTreatment` rewritten: control types, `UserForm_Initialize`, new `txtCustomerName_Exit` behavior, new `txtStart_Change`/`txtEnd_Change`, `SaveTreatmentForm`'s `CreateTreatmentRecord` call. 3 new Public helper Subs added: `LoadTherapistList`, `RefreshCardList`, `RecalcSessions` |
| `modCard.bas` | `FindUsableCard`'s eligibility test factored into a shared `CardEligible` helper; new `ListEligibleCards` function added (uses the same helper — see Section 7) |
| `modSystemCore.bas` | New `FindHeaderColumn` helper (header-based column lookup, used for `Therapist_Master`) |
| `modFormLabels.bas` | One-character-code typo fix: 療療類型 -> 治療類型 |

## 5. Unmodified Modules (13 of 17)

`modAudit.bas`, `modCardUsage.bas`, `modCustomer.bas`, `modDashboard.bas`,
`modErrorHandler.bas`, `modInventory.bas`, `modM5_QA.bas`, `modMainMenu.bas`,
`modProduct.bas`, `modSecurity.bas`, `modText.bas`, `modTreatment.bas`,
`modValidation.bas` — byte-for-byte identical to the confirmed Hotfix-01 baseline.
**`modTreatment.CreateTreatmentRecord` and its 22-column write are completely
untouched** — the fix only changes what values the form passes in, never the
transaction core itself.

## 6. Control Manifest Changes (frmTreatment only; all 7 other forms untouched)

| Control | Before | After | Notes |
|---|---|---|---|
| `cboTherapist` | ComboBox, empty | ComboBox, 2 columns (Name shown, ID hidden), populated from `Therapist_Master` | Name unchanged |
| `txtTreatmentType` | TextBox | **ComboBox** (Name unchanged), 2 fixed items | Per GPT ruling #1 |
| `txtBookingType` | TextBox | **ComboBox** (Name unchanged), 4 fixed items | Per GPT ruling #1 |
| `cboCardID` | ComboBox, empty | ComboBox, repopulated on `txtCustomerName_Exit` | Name unchanged |
| `txtStart` / `txtEnd` | no event handlers | `_Change` handlers added, call shared `RecalcSessions` | New behavior, no control change |

No control was renamed. No control was removed. `cboPayment`, `cboShortage`,
`txtCustomerName`, `txtCustomerID`, `txtTreatmentDate`, `txtSessions`, `txtRevenue`,
`txtShortageAmount`, `txtNotes`, `cmdSave`, `cmdCancel` are all unchanged.

## 7. DataSource Detail

- **Therapist**: `Therapist_Master`, columns located by header (`FindHeaderColumn`) for
  `TherapistID`/`TherapistName`/`Status`. A row is loaded if `Status` is blank, `"Active"`
  (case-insensitive), or `"在職"` — both English and Chinese status values are accepted
  because the actual value convention used on the live sheet was not independently
  verifiable by Claude before this hotfix (see Section 12). This is a deliberately
  permissive filter, not a guess baked into a single hard-coded comparison: if `Status`
  turns out to use a different convention Arthur/GPT did not mention, the safer failure
  mode is "shows an inactive therapist" rather than "silently shows nobody."
- **Treatment Type**: fixed list, GPT ruling #2 (一般調理, 體驗調理).
- **Booking Type**: fixed list, GPT ruling #2 (預約, Walk-in, 改期, 店長安排).
- **CardID**: `Card_Master`, via new `modCard.ListEligibleCards`, which shares its
  eligibility test (`CardEligible`) with the existing `FindUsableCard` used by the
  transaction core — not a second, independently-defined rule. `CardStatus` is **not**
  filtered (unchanged limitation, carried forward from the original PATCH-02 report —
  see Section 12).

## 8. Session Calculation

Reuses `modTreatment.CalculateSessions(startTime, endTime)` exactly as it already
existed (the 40-minute rule; 10:00-11:20 = 2). Not reimplemented, not duplicated.
Wired to `txtStart_Change`/`txtEnd_Change` via the new `RecalcSessions` helper, which:
- does nothing if either field is not a valid time (blank or malformed input)
- does nothing if End <= Start
- otherwise sets `txtSessions.Value` to the calculated result

`txtSessions` remains directly editable by the user (unchanged) — auto-calculation is a
convenience, not a lock.

## 9. Card Filter

`modCard.ListEligibleCards(customerID, treatmentDate)` returns every `Card_Master` row
where `CustomerID` matches, `RemainingSessions > 0`, and `PurchaseDate <= treatmentDate`
— identical conditions to `FindUsableCard`, sorted oldest-`PurchaseDate`-first. Triggered
from `txtCustomerName_Exit`, immediately after `CustomerID` is resolved, per the task's
explicit requirement ("CustomerID 改變後，CardID 下拉必須重新載入").

## 10. Static QA

Performed on all 17 files in this delivery, not just the 4 changed ones:

| Check | Result |
|---|---|
| ASCII-safe source (17/17) | PASS |
| Sub/Function begin-end balance (17/17) | PASS |
| Duplicate procedure names across files | PASS (none) |
| Cross-module reference (parenthesized calls, all 17 files) | PASS — 4 initial hits were array-element access (`headers(i)`, `cp(i)`, `ids(n)`, `dates(n)`), confirmed false positives by inspection, not genuine calls |
| Cross-module reference (paren-less Sub-style calls, continuations joined, all 17 files) | PASS (none) |
| Max `_` line-continuation chain in `modFormBuilder.bas` | 2 (unchanged from Hotfix-01, still far under the 24 limit) |
| Longest physical line | 443 characters (limit 1023) |
| Control name vs. `Me.<control>` handler cross-reference, all 8 forms | PASS |
| Generated `frmTreatment` code reconstructed and read in full | Done — see Section 13 for the actual text |

**Static QA verdict: PASS.**

## 11. Not Claimed

Excel Compile PASS, Runtime PASS, UAT PASS, Production Ready — none of these are
claimed. All remain `UNVERIFIED`.

## 12. Known Limitations / Not Fixed (explicitly out of Hotfix-02 scope)

- `CardStatus` is still not used as an eligibility filter in `ListEligibleCards`/
  `FindUsableCard` — this is the same limitation already reported in the original
  PATCH-02 Implementation Report (2026-09-22), not newly introduced or newly
  discovered. Per the task's own instruction ("若尚未完成，於 Report 標示，不得自行
  擴大範圍"), it is reported again here rather than fixed.
- Therapist `Status` value convention (English `"Active"` vs Chinese `"在職"`) was not
  independently confirmed by Claude before this hotfix — the dual-accept filter in
  Section 7 is a deliberate defensive choice, not a verified fact. Arthur's TEST-05
  re-run will confirm which (or both) actually appear.
- CardID dropdown is refreshed only when `CustomerID` changes (via
  `txtCustomerName_Exit`), exactly as the task specified. It does **not** also refresh
  when `TreatmentDate` is changed afterward — the task's requirement only named
  CustomerID as the trigger, so this was not added, to avoid expanding scope beyond
  what was approved.

## 13. Reconstructed Generated Code (for review, not the source itself)

The actual VBA text that `BuildTreatment` injects into `frmTreatment`'s code module was
mechanically reconstructed from the `AppendCode` call sequence in `modFormBuilder.bas`
and is reproduced in full below, so this can be read without a VBA environment:

```vb
Private Sub UserForm_Initialize()
    Me.txtTreatmentDate.Value = Format$(Date, "yyyy-mm-dd")
    Me.cboPayment.AddItem TW_Cash()
    Me.cboPayment.AddItem TW_Bank()
    Me.cboShortage.AddItem ""
    Me.cboShortage.AddItem TW_Charge()
    Me.cboShortage.AddItem TW_Waive()
    Me.cboShortage.AddItem TW_Pending()
    Me.txtTreatmentType.AddItem UText(&H4E00&, &H822C&, &H8ABF&, &H7406&)   ' 一般調理
    Me.txtTreatmentType.AddItem UText(&H9AD4&, &H9A57&, &H8ABF&, &H7406&)   ' 體驗調理
    Me.txtBookingType.AddItem UText(&H9810&, &H7D04&)                       ' 預約
    Me.txtBookingType.AddItem "Walk-in"
    Me.txtBookingType.AddItem UText(&H6539&, &H671F&)                       ' 改期
    Me.txtBookingType.AddItem UText(&H5E97&, &H9577&, &H5B89&, &H6392&)     ' 店長安排
    Me.cboTherapist.ColumnCount = 2
    Me.cboTherapist.ColumnWidths = "100 pt;0 pt"
    LoadTherapistList Me.cboTherapist
End Sub
Private Sub txtCustomerName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Me.txtCustomerID.Value = GetCustomerByName(Me.txtCustomerName.Value)
    RefreshCardList Me.cboCardID, Me.txtCustomerID.Value, Me.txtTreatmentDate.Value
End Sub
Private Sub txtStart_Change()
    RecalcSessions Me.txtStart, Me.txtEnd, Me.txtSessions
End Sub
Private Sub txtEnd_Change()
    RecalcSessions Me.txtStart, Me.txtEnd, Me.txtSessions
End Sub
Public Sub SaveTreatmentForm()
    Dim rid As String, res As String, pay As String
    On Error GoTo EH
    If Me.cboPayment.ListIndex = 0 Then pay = TW_Cash() Else pay = TW_Bank()
    If Me.cboShortage.ListIndex > 0 Then res = Me.cboShortage.Value Else res = ""
    rid = CreateTreatmentRecord(CDate(Me.txtTreatmentDate.Value), TimeValue(Me.txtStart.Value), TimeValue(Me.txtEnd.Value), Me.txtCustomerName.Value, Me.cboTherapist.Value, Me.cboTherapist.Column(1), CDbl(Me.txtSessions.Value), Me.txtTreatmentType.Value, Me.txtBookingType.Value, Me.cboCardID.Value, pay, CDbl(Val(Me.txtRevenue.Value)), "", Me.txtNotes.Value, res, CDbl(Val(Me.txtShortageAmount.Value)))
    MsgBox "Saved: " & rid, vbInformation, "NPSBS IOMS"
    Unload Me
    Exit Sub
EH:
    MsgBox Err.Description, vbExclamation, "NPSBS IOMS"
End Sub
```

`CreateTreatmentRecord`'s call now carries 16 arguments, matching its unchanged
signature (9 required + 7 optional). Argument 6 (`therapistID`) changed from a hard-coded
empty string to `Me.cboTherapist.Column(1)` — this was a pre-existing gap (TherapistID
was always blank regardless of the TEST-05 failures) that Claude closed as a direct,
minimal completion of the task's own "保留 Therapist/TherapistID 現有架構" requirement,
using data already being read for the dropdown fix — flagged here explicitly rather than
silently bundled in.

## 14. SHA-256 Manifest

See companion file `MANIFEST_SHA256.txt` (all 17 files, so Arthur/Gemini can confirm
which 4 actually changed vs. the Hotfix-01 baseline by comparing hashes).

## 15. Git Commit

Not applicable — Claude has no GitHub write access in this environment (established
earlier in this collaboration). Arthur uploads the delivered files manually, as with
every prior round.

## 16. Next Required Action

```
Claude (this hotfix)
  -> Gemini Independent QA (see task Section 9 checklist)
  -> GPT Architecture Gate
  -> Arthur Excel TEST-05 re-run
  -> GPT Final Gate
```
