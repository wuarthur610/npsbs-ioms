Attribute VB_Name = "modTreatment"
Option Explicit

Public Function CalculateSessions(ByVal startTime As Date, ByVal endTime As Date) As Double
    If endTime<startTime Then Err.Raise vbObjectError+1401,"IOMS",MsgText("END_BEFORE_START")
    CalculateSessions=Round(((endTime-startTime)*1440#)/40#,2)
End Function

Public Function CalculateFinalSessions(ByVal startTime As Date, ByVal endTime As Date, ByVal adjustedSessions As Variant) As Double
    If Len(Trim$(CStr(Nz(adjustedSessions,""))))>0 Then CalculateFinalSessions=SafeNumber(adjustedSessions) Else CalculateFinalSessions=CalculateSessions(startTime,endTime)
End Function

' ============================================================
' PATCH-02 / P0-01, P0-02: Production transaction entry point.
' Name UNCHANGED (CreateTreatmentRecord) per Contract V1.1 S16 - this
' function remains the ONLY Production Treatment transaction API.
'
' Flow (Contract V1.1 S4/S5/S6):
'   Phase 1 - VALIDATE (pure reads only, zero mutation):
'     Validate Customer -> Validate Therapist/Treatment -> Validate Payment
'     -> Resolve Card -> Calculate Deduction (Balance Guard enforced here)
'     -> Validate Shortage Resolution (+ Shortage Amount if Charge)
'   Phase 2 - MUTATE (compensatable pseudo-transaction):
'     Create Treatment_Record -> Commit Deduction -> Create Card_Usage_Detail
'     -> Handle Shortage -> Audit -> COMMIT
'   On failure during Phase 2: Treatment_Record.Status -> "Voided" (never
'   deleted), Card_Master restored to its pre-transaction snapshot, any
'   Card_Usage_Detail row already created is compensated with a Rollback
'   row. A failure during Phase 1 has nothing to compensate, since nothing
'   was written yet - this is exactly what P0-01 is for.
' ============================================================
Public Function CreateTreatmentRecord(ByVal treatmentDate As Date, ByVal startTime As Date, ByVal endTime As Date, ByVal customerName As String, ByVal therapistName As String, ByVal therapistID As String, ByVal sessions As Double, ByVal treatmentType As String, ByVal bookingType As String, Optional ByVal cardID As String="", Optional ByVal paymentMethod As String="", Optional ByVal treatmentRevenue As Double=0, Optional ByVal calendarUID As String="", Optional ByVal notes As String="", Optional ByVal shortageResolution As String="", Optional ByVal shortageAmount As Double=0) As String
    Dim ws As Worksheet, r As Long, id As String, custID As String
    Dim actualCard As String, status As String
    Dim beforeUsed As Double, beforeRemain As Double, usedS As Double, shortage As Double, afterRemain As Double
    Dim recordCreated As Boolean, cardCommitted As Boolean, usageID As String, usageCreated As Boolean, pendingID As String

    EnsureProductionSheets

    ' ---------------- Phase 1: VALIDATE (no mutation) ----------------
    custID=GetCustomerByName(customerName)
    If custID="" Then Err.Raise vbObjectError+1402,"IOMS",MsgText("NO_CUSTOMER")
    ValidateTreatment customerName,therapistName,treatmentDate,sessions
    ValidatePaymentMethod paymentMethod

    If Len(cardID)=0 Then
        actualCard=FindUsableCard(custID,treatmentDate)
    Else
        actualCard=cardID
        ValidateCardForCustomer actualCard,custID,treatmentDate
    End If

    beforeUsed=0: beforeRemain=0: usedS=0: shortage=0: afterRemain=0
    If Len(actualCard)>0 Then
        CalculateDeduction actualCard,sessions,beforeUsed,beforeRemain,usedS,shortage,afterRemain
    End If

    If Len(actualCard)=0 Then
        status="NoCard"
    ElseIf shortage>0 Then
        status="Insufficient"
        ' Contract V1.1 S4 / PATCH_SPEC P0-01: shortage resolution (and the
        ' Charge amount, if applicable) MUST be validated before any
        ' persistent mutation - this is the exact defect PATCH-02 fixes.
        ValidateShortageResolution shortage,shortageResolution
        If NormalizeResolution(shortageResolution)="CHARGE" Then
            If shortageAmount<=0 Then Err.Raise vbObjectError+1201,"IOMS",MsgText("SHORTAGE_FEE")
        End If
    Else
        status="AutoDeducted"
    End If

    ' ---------------- Phase 2: MUTATE (compensatable from here on) ----------------
    On Error GoTo Fail

    Set ws=ThisWorkbook.Worksheets("Treatment_Record")
    r=LastDataRow(ws,1)+1: id=NextID("TRT-",ws,1)
    ws.Cells(r,1)=id: ws.Cells(r,2)=treatmentDate: ws.Cells(r,3)=startTime: ws.Cells(r,4)=endTime: ws.Cells(r,5)=customerName: ws.Cells(r,6)=custID
    ws.Cells(r,7)=therapistName: ws.Cells(r,8)=therapistID: ws.Cells(r,9)=sessions: ws.Cells(r,10)=treatmentType: ws.Cells(r,11)=bookingType: ws.Cells(r,12)=actualCard
    ws.Cells(r,13)=status: ws.Cells(r,14)=shortageResolution: ws.Cells(r,15)=shortage: ws.Cells(r,16)=treatmentRevenue: ws.Cells(r,17)=paymentMethod: ws.Cells(r,18)=calendarUID: ws.Cells(r,19)=notes: ws.Cells(r,20)=Now: ws.Cells(r,21)=CurrentUser(): ws.Cells(r,22)="Completed"
    recordCreated=True

    If Len(actualCard)>0 And usedS>0 Then
        CommitDeduction actualCard,beforeRemain,usedS,afterRemain
        cardCommitted=True
        usageID=CreateUsageDetail(id,actualCard,custID,treatmentDate,beforeRemain,usedS,afterRemain,"TreatmentDeduction",shortage,notes)
        usageCreated=(Len(usageID)>0)
    End If

    If shortage>0 Then
        pendingID=HandleShortage(id,custID,customerName,shortage,shortageResolution,shortageAmount,notes)
    End If

    WriteAuditLog "CREATE","Treatment_Record",id,"Status","","Completed"
    CreateTreatmentRecord=id
    Exit Function

Fail:
    Dim failNum As Long, failDesc As String
    failNum=Err.Number: failDesc=Err.Description
    On Error Resume Next   ' compensation itself must not throw us out of the handler
    If usageCreated Then CompensateUsageDetail usageID,"Treatment transaction failed: " & failDesc
    If cardCommitted Then RestoreCardSnapshot actualCard,beforeUsed,beforeRemain,"Treatment transaction failed: " & failDesc
    If recordCreated Then
        ws.Cells(r,22)="Voided"
        WriteAuditLog "COMPENSATE","Treatment_Record",id,"Status","Completed","Voided","Transaction failed: " & failDesc
    End If
    On Error GoTo 0
    Err.Raise failNum,"IOMS",failDesc
End Function

Public Function SessionAdjusted(ByVal adjustedSessions As Variant) As Boolean
    SessionAdjusted=(Len(Trim$(CStr(Nz(adjustedSessions,""))))>0)
End Function
