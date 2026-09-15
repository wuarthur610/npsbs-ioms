Attribute VB_Name = "modInventory"
Option Explicit

Public Function CalculateBookQuantity(ByVal productCode As String) As Double
    Dim ws As Worksheet, r As Long
    On Error Resume Next: Set ws=ThisWorkbook.Worksheets("Inventory"): On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r=LastDataRow(ws,1) To 2 Step -1: If CStr(ws.Cells(r,3).Value)=productCode And CStr(ws.Cells(r,8).Value)="Closed" Then CalculateBookQuantity=SafeNumber(ws.Cells(r,5).Value): Exit Function
    Next r
End Function

Public Function CalculateVariance(ByVal bookQty As Double, ByVal actualQty As Double) As Double
    CalculateVariance=actualQty-bookQty
End Function

Public Function ProcessCountVariance(ByVal productCode As String, ByVal countDate As Date, ByVal actualQty As Double, ByVal reason As String, Optional ByVal notes As String="") As String
    Dim ws As Worksheet, r As Long, id As String, bookQty As Double, variance As Double
    EnsureProductionSheets: Set ws=ThisWorkbook.Worksheets("Inventory")
    bookQty=CalculateBookQuantity(productCode): variance=CalculateVariance(bookQty,actualQty): ValidateVarianceReason variance,reason
    If variance<>0 And reason="" Then Err.Raise vbObjectError+1501,"IOMS",MsgText("INV_REASON")
    r=LastDataRow(ws,1)+1: id=NextID("INV-",ws,1)
    ws.Cells(r,1)=id: ws.Cells(r,2)=countDate: ws.Cells(r,3)=productCode: ws.Cells(r,4)=bookQty: ws.Cells(r,5)=actualQty: ws.Cells(r,6)=variance: ws.Cells(r,7)=reason: ws.Cells(r,8)="Closed": ws.Cells(r,9)=notes: ws.Cells(r,10)=Now: ws.Cells(r,11)=CurrentUser()
    WriteAuditLog "CREATE","Inventory",id,"Variance",CStr(variance),"0","Reconciled: " & reason
    ProcessCountVariance=id
End Function
