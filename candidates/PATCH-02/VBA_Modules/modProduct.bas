Attribute VB_Name = "modProduct"
Option Explicit

Public Function GetProduct(ByVal productCode As String) As String
    Dim ws As Worksheet, r As Long
    On Error Resume Next: Set ws=ThisWorkbook.Worksheets("Product_Master"): On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r=2 To LastDataRow(ws,1): If CStr(ws.Cells(r,1).Value)=productCode Then GetProduct=CStr(ws.Cells(r,2).Value): Exit Function
    Next r
End Function

Public Function SaveCardProductVersion(ByVal cardID As String, ByVal productCode As String, ByVal productName As String, ByVal quantity As Double, ByVal unitPrice As Double, Optional ByVal notes As String="") As String
    Dim ws As Worksheet, r As Long, id As String, nextVer As Long, v As Variant
    Set ws=EnsureSheet("Card_Product_Detail",Array("DetailID","CardID","VersionNo","ProductCode","ProductName","Quantity","UnitPrice","LineAmount","EffectiveFrom","EffectiveTo","Notes","CreatedAt","CreatedBy"))
    For r=2 To LastDataRow(ws,1): If CStr(ws.Cells(r,2).Value)=cardID Then v=SafeNumber(ws.Cells(r,3).Value): If v>=nextVer Then nextVer=v+1
    Next r
    If nextVer=0 Then nextVer=1
    r=LastDataRow(ws,1)+1: id=NextID("CPD-",ws,1)
    ws.Cells(r,1)=id: ws.Cells(r,2)=cardID: ws.Cells(r,3)=nextVer: ws.Cells(r,4)=productCode: ws.Cells(r,5)=productName: ws.Cells(r,6)=quantity: ws.Cells(r,7)=unitPrice: ws.Cells(r,8)=quantity*unitPrice: ws.Cells(r,9)=Date: ws.Cells(r,11)=notes: ws.Cells(r,12)=Now: ws.Cells(r,13)=CurrentUser()
    SaveCardProductVersion=id
End Function
