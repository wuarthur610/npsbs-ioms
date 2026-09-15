Attribute VB_Name = "modTreatment"
Option Explicit

Public Function CalculateSessions(ByVal startTime As Date, ByVal endTime As Date) As Double
    If endTime<startTime Then Err.Raise vbObjectError+1401,"IOMS",MsgText("END_BEFORE_START")
    CalculateSessions=Round(((endTime-startTime)*1440#)/40#,2)
End Function

Public Function CalculateFinalSessions(ByVal startTime As Date, ByVal endTime As Date, ByVal adjustedSessions As Variant) As Double
    If Len(Trim$(CStr(Nz(adjustedSessions,""))))>0 Then CalculateFinalSessions=SafeNumber(adjustedSessions) Else CalculateFinalSessions=CalculateSessions(startTime,endTime)
End Function

Public Function CreateTreatmentRecord(ByVal treatmentDate As Date, ByVal startTime As Date, ByVal endTime As Date, ByVal customerName As String, ByVal therapistName As String, ByVal therapistID As String, ByVal sessions As Double, ByVal treatmentType As String, ByVal bookingType As String, Optional ByVal cardID As String="", Optional ByVal paymentMethod As String="", Optional ByVal treatmentRevenue As Double=0, Optional ByVal calendarUID As String="", Optional ByVal notes As String="", Optional ByVal shortageResolution As String="", Optional ByVal shortageAmount As Double=0) As String
    Dim ws As Worksheet, r As Long, id As String, custID As String, shortage As Double, actualCard As String, used As Double, status As String, pendingID As String
    On Error GoTo EH
    EnsureProductionSheets
    custID=GetCustomerByName(customerName)
    If custID="" Then Err.Raise vbObjectError+1402,"IOMS",MsgText("NO_CUSTOMER")
    ValidateTreatment customerName,therapistName,treatmentDate,sessions: ValidatePaymentMethod paymentMethod
    If Len(cardID)=0 Then actualCard=FindUsableCard(custID,treatmentDate) Else actualCard=cardID: ValidateCardForCustomer actualCard,custID,treatmentDate
    used=0: shortage=0
    If Len(actualCard)>0 Then If DeductSessions(actualCard,sessions,shortage) Then used=sessions-shortage
    If Len(actualCard)=0 Then status="NoCard" Else If shortage>0 Then status="Insufficient" Else status="AutoDeducted"
    Set ws=ThisWorkbook.Worksheets("Treatment_Record"): r=LastDataRow(ws,1)+1: id=NextID("TRT-",ws,1)
    ws.Cells(r,1)=id: ws.Cells(r,2)=treatmentDate: ws.Cells(r,3)=startTime: ws.Cells(r,4)=endTime: ws.Cells(r,5)=customerName: ws.Cells(r,6)=custID
    ws.Cells(r,7)=therapistName: ws.Cells(r,8)=therapistID: ws.Cells(r,9)=sessions: ws.Cells(r,10)=treatmentType: ws.Cells(r,11)=bookingType: ws.Cells(r,12)=actualCard
    ws.Cells(r,13)=status: ws.Cells(r,14)=shortageResolution: ws.Cells(r,15)=shortage: ws.Cells(r,16)=treatmentRevenue: ws.Cells(r,17)=paymentMethod: ws.Cells(r,18)=calendarUID: ws.Cells(r,19)=notes: ws.Cells(r,20)=Now: ws.Cells(r,21)=CurrentUser(): ws.Cells(r,22)="Completed"
    If used>0 Then CreateUsageDetail id,actualCard,custID,treatmentDate,used,"TreatmentDeduction",shortage,notes
    If shortage>0 Then pendingID=HandleShortage(id,custID,customerName,shortage,shortageResolution,shortageAmount,notes)
    WriteAuditLog "CREATE","Treatment_Record",id,"Status","","Completed"
    CreateTreatmentRecord=id: Exit Function
EH:
    Err.Raise Err.Number,"IOMS",Err.Description
End Function

Public Function SessionAdjusted(ByVal adjustedSessions As Variant) As Boolean
    SessionAdjusted=(Len(Trim$(CStr(Nz(adjustedSessions,""))))>0)
End Function
