Attribute VB_Name = "modCardUsage"
Option Explicit

Public Function CreateUsageDetail(ByVal treatmentRecordID As String, ByVal cardID As String, ByVal customerID As String, ByVal usageDate As Date, ByVal sessionsUsed As Double, ByVal usageType As String, Optional ByVal shortageSessions As Double=0, Optional ByVal notes As String="") As String
    Dim ws As Worksheet, r As Long, id As String
    If Len(cardID)=0 Or sessionsUsed<=0 Then Exit Function
    Set ws=EnsureSheet("Card_Usage_Detail",Array("UsageDetailID","TreatmentRecordID","CardID","CustomerID","UsageDate","SessionsUsed","UsageType","ShortageSessions","Notes","CreatedAt","CreatedBy","Version"))
    r=LastDataRow(ws,1)+1: id=NextID("USE-",ws,1)
    ws.Cells(r,1)=id: ws.Cells(r,2)=treatmentRecordID: ws.Cells(r,3)=cardID: ws.Cells(r,4)=customerID: ws.Cells(r,5)=usageDate
    ws.Cells(r,6)=sessionsUsed: ws.Cells(r,7)=usageType: ws.Cells(r,8)=shortageSessions: ws.Cells(r,9)=notes: ws.Cells(r,10)=Now: ws.Cells(r,11)=CurrentUser(): ws.Cells(r,12)=1
    CreateUsageDetail=id
End Function

Public Sub LinkTreatmentRecord(ByVal usageDetailID As String, ByVal treatmentRecordID As String)
    Dim ws As Worksheet, r As Long
    Set ws=ThisWorkbook.Worksheets("Card_Usage_Detail")
    For r=2 To LastDataRow(ws,1): If CStr(ws.Cells(r,1).Value)=usageDetailID Then ws.Cells(r,2)=treatmentRecordID: Exit Sub
    Next r
End Sub

Public Sub SettlePending(ByVal pendingID As String, ByVal newCardID As String, ByVal sessions As Double)
    Dim wp As Worksheet, wc As Worksheet, r As Long, pr As Long, cid As String, custID As String, tid As String, oldRemain As Double, useN As Double, newUseID As String
    If sessions<=0 Then Err.Raise vbObjectError+1301,"IOMS",MsgText("SESSION_REQUIRED")
    Set wp=ThisWorkbook.Worksheets("Pending_Shortage"): Set wc=ThisWorkbook.Worksheets("Card_Master")
    For r=2 To LastDataRow(wp,1)
        If CStr(wp.Cells(r,1).Value)=pendingID Then
            If CStr(wp.Cells(r,7).Value)<>"Pending" Then Err.Raise vbObjectError+1302,"IOMS",MsgText("PENDING_CLOSED")
            tid=CStr(wp.Cells(r,2).Value): custID=CStr(wp.Cells(r,3).Value): ValidateCardForCustomer newCardID,custID,Date
            If sessions>SafeCardBalance(wc,newCardID) Then Err.Raise vbObjectError+1304,"IOMS",MsgText("NEW_CARD_SHORT")
            useN=sessions
            newUseID=CreateUsageDetail(tid,newCardID,custID,Date,useN,"PendingSettlement",0,"Pending settlement; original Treatment_Record unchanged")
            DeductSessions newCardID,useN,oldRemain
            wp.Cells(r,7)="Settled": wp.Cells(r,8)=Date: wp.Cells(r,9)=newCardID
            WriteAuditLog "UPDATE","Pending_Shortage",pendingID,"Status","Pending","Settled","New card settlement " & newUseID
            Exit Sub
        End If
    Next r
    Err.Raise vbObjectError+1305,"IOMS",MsgText("PENDING_NOT_FOUND")
End Sub

Private Function SafeCardBalance(ByVal ws As Worksheet, ByVal cardID As String) As Double
    Dim r As Long
    For r=2 To LastDataRow(ws,1): If CStr(ws.Cells(r,1).Value)=cardID Then SafeCardBalance=SafeNumber(ws.Cells(r,7).Value): Exit Function
    Next r
End Function
