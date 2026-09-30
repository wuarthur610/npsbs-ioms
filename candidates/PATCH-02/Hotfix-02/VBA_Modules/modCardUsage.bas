Attribute VB_Name = "modCardUsage"
Option Explicit

' PATCH-02 / P0-06: added BeforeSessions and AfterSessions as REQUIRED
' parameters (no silent zero-fill), inserted around SessionsUsed so a caller
' cannot forget them. Physical column layout: existing 1-12 unchanged,
' BeforeSessions/AfterSessions appended at 13/14 (Contract V1.1 S11).
' Caller must pass the SAME beforeSessions/afterSessions values that were
' produced by modCard.CalculateDeduction / CommitDeduction in the SAME
' transaction - this function never recomputes them.
Public Function CreateUsageDetail(ByVal treatmentRecordID As String, ByVal cardID As String, _
        ByVal customerID As String, ByVal usageDate As Date, _
        ByVal beforeSessions As Double, ByVal sessionsUsed As Double, ByVal afterSessions As Double, _
        ByVal usageType As String, Optional ByVal shortageSessions As Double = 0, _
        Optional ByVal notes As String = "") As String
    Dim ws As Worksheet, r As Long, id As String
    If Len(cardID)=0 Or sessionsUsed<=0 Then Exit Function
    Set ws=EnsureSheet("Card_Usage_Detail",Array("UsageDetailID","TreatmentRecordID","CardID","CustomerID","UsageDate","SessionsUsed","UsageType","ShortageSessions","Notes","CreatedAt","CreatedBy","Version","BeforeSessions","AfterSessions"))
    EnsureColumnHeader ws, 13, "BeforeSessions"
    EnsureColumnHeader ws, 14, "AfterSessions"
    r=LastDataRow(ws,1)+1: id=NextID("USE-",ws,1)
    ws.Cells(r,1)=id: ws.Cells(r,2)=treatmentRecordID: ws.Cells(r,3)=cardID: ws.Cells(r,4)=customerID: ws.Cells(r,5)=usageDate
    ws.Cells(r,6)=sessionsUsed: ws.Cells(r,7)=usageType: ws.Cells(r,8)=shortageSessions: ws.Cells(r,9)=notes: ws.Cells(r,10)=Now: ws.Cells(r,11)=CurrentUser(): ws.Cells(r,12)=1
    ws.Cells(r,13)=beforeSessions: ws.Cells(r,14)=afterSessions
    CreateUsageDetail=id
End Function

Public Sub LinkTreatmentRecord(ByVal usageDetailID As String, ByVal treatmentRecordID As String)
    Dim ws As Worksheet, r As Long
    Set ws=ThisWorkbook.Worksheets("Card_Usage_Detail")
    For r=2 To LastDataRow(ws,1): If CStr(ws.Cells(r,1).Value)=usageDetailID Then ws.Cells(r,2)=treatmentRecordID: Exit Sub
    Next r
End Sub

' PATCH-02 / P0-06: compensating entry for a previously created usage detail
' row. Creates a NEW row (append-only, never edits the original) with
' UsageType = "Rollback", reversing Before/After so the compensation reads
' naturally alongside the original entry it cancels out.
Public Function CompensateUsageDetail(ByVal originalUsageDetailID As String, ByVal reason As String) As String
    Dim ws As Worksheet, r As Long
    Dim origTID As String, origCard As String, origCust As String
    Dim origUsed As Double, origBefore As Double, origAfter As Double, foundRow As Boolean

    Set ws=ThisWorkbook.Worksheets("Card_Usage_Detail")
    For r=2 To LastDataRow(ws,1)
        If CStr(ws.Cells(r,1).Value)=originalUsageDetailID Then
            origTID=CStr(ws.Cells(r,2).Value): origCard=CStr(ws.Cells(r,3).Value): origCust=CStr(ws.Cells(r,4).Value)
            origUsed=SafeNumber(ws.Cells(r,6).Value)
            origBefore=SafeNumber(ws.Cells(r,13).Value): origAfter=SafeNumber(ws.Cells(r,14).Value)
            foundRow=True
            Exit For
        End If
    Next r
    If Not foundRow Then Exit Function

    CompensateUsageDetail=CreateUsageDetail(origTID, origCard, origCust, Date, origAfter, origUsed, origBefore, _
        "Rollback", 0, "Compensates " & originalUsageDetailID & ": " & reason)
End Function

' ============================================================
' PATCH-02 / P0-05: Pending Settlement Atomicity.
' Order is: Validate -> Snapshot (pure calc, no mutation) -> Commit deduction
' -> Create settlement usage detail -> Mark Pending_Shortage Settled -> Audit.
' On any failure after CommitDeduction, the card is restored to its
' pre-settlement snapshot and (if it was created) the usage detail is
' compensated with a Rollback row. Pending_Shortage is left "Pending" on
' failure - it is never marked Settled unless every step above succeeded.
' ============================================================
Public Sub SettlePending(ByVal pendingID As String, ByVal newCardID As String, ByVal sessions As Double)
    Dim wp As Worksheet, r As Long, custID As String, tid As String
    Dim beforeUsed As Double, beforeRemain As Double, usedS As Double, shortS As Double, afterRemain As Double
    Dim cardCommitted As Boolean, usageID As String, usageCreated As Boolean

    If sessions<=0 Then Err.Raise vbObjectError+1301,"IOMS",MsgText("SESSION_REQUIRED")

    Set wp=ThisWorkbook.Worksheets("Pending_Shortage")
    r=0
    Dim rr As Long
    For rr=2 To LastDataRow(wp,1)
        If CStr(wp.Cells(rr,1).Value)=pendingID Then r=rr: Exit For
    Next rr
    If r=0 Then Err.Raise vbObjectError+1305,"IOMS",MsgText("PENDING_NOT_FOUND")
    If CStr(wp.Cells(r,7).Value)<>"Pending" Then Err.Raise vbObjectError+1302,"IOMS",MsgText("PENDING_CLOSED")

    tid=CStr(wp.Cells(r,2).Value): custID=CStr(wp.Cells(r,3).Value)

    ' ---------------- Phase 1: VALIDATE (no mutation) ----------------
    ValidateCardForCustomer newCardID, custID, Date
    If Not CalculateDeduction(newCardID, sessions, beforeUsed, beforeRemain, usedS, shortS, afterRemain) Then
        Err.Raise vbObjectError+1303,"IOMS",MsgText("NO_CARD")
    End If
    If shortS>0 Then Err.Raise vbObjectError+1304,"IOMS",MsgText("NEW_CARD_SHORT")

    ' ---------------- Phase 2: MUTATE (compensatable from here on) ----------------
    On Error GoTo Fail

    CommitDeduction newCardID, beforeRemain, usedS, afterRemain
    cardCommitted=True

    usageID=CreateUsageDetail(tid, newCardID, custID, Date, beforeRemain, usedS, afterRemain, _
        "PendingSettlement", 0, "Pending settlement; original Treatment_Record unchanged")
    usageCreated=(Len(usageID)>0)

    wp.Cells(r,7)="Settled": wp.Cells(r,8)=Date: wp.Cells(r,9)=newCardID
    WriteAuditLog "UPDATE","Pending_Shortage",pendingID,"Status","Pending","Settled","New card settlement " & usageID
    Exit Sub

Fail:
    Dim failNum As Long, failDesc As String
    failNum=Err.Number: failDesc=Err.Description
    On Error Resume Next
    If usageCreated Then CompensateUsageDetail usageID, "Pending settlement failed: " & failDesc
    If cardCommitted Then RestoreCardSnapshot newCardID, beforeUsed, beforeRemain, "Pending settlement failed: " & failDesc
    WriteAuditLog "COMPENSATE","Pending_Shortage",pendingID,"Status","Pending","Pending","Settlement failed and compensated: " & failDesc
    On Error GoTo 0
    Err.Raise failNum,"IOMS", MsgText("SETTLEMENT_FAILED") & " " & failDesc
End Sub
