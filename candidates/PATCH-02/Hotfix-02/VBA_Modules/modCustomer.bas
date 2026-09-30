Attribute VB_Name = "modCustomer"
Option Explicit

Public Function GetCustomerByName(ByVal customerName As String) As String
    Dim ws As Worksheet, r As Long, lastRow As Long, found As String
    Set ws=ThisWorkbook.Worksheets("Customer_Master")
    lastRow=LastDataRow(ws,2)
    For r=2 To lastRow
        If Trim$(CStr(ws.Cells(r,2).Value))=Trim$(customerName) Then
            If found<>"" Then Err.Raise vbObjectError+1001,"IOMS",MsgText("DUP_CUSTOMER")
            found=CStr(ws.Cells(r,1).Value)
        End If
    Next r
    GetCustomerByName=found
End Function

Public Function CreateCustomer(ByVal customerName As String, Optional ByVal notes As String="") As String
    Dim ws As Worksheet, r As Long, id As String
    RequireValue customerName,MsgText("CUSTOMER_REQUIRED")
    If GetCustomerByName(customerName)<>"" Then Err.Raise vbObjectError+1002,"IOMS",MsgText("CUSTOMER_EXISTS")
    Set ws=ThisWorkbook.Worksheets("Customer_Master"): r=LastDataRow(ws,1)+1
    id=NextID("CUST-",ws,1)
    ws.Cells(r,1)=id: ws.Cells(r,2)=Trim$(customerName): ws.Cells(r,3)=Date: ws.Cells(r,4)=notes: ws.Cells(r,5)="Active"
    WriteAuditLog "CREATE","Customer_Master",id,"CustomerName","",customerName
    CreateCustomer=id
End Function

Public Sub UpdateCustomerID(ByVal customerName As String, ByVal newCustomerID As String)
    Dim ws As Worksheet, r As Long, oldID As String
    Set ws=ThisWorkbook.Worksheets("Customer_Master")
    For r=2 To LastDataRow(ws,2)
        If Trim$(CStr(ws.Cells(r,2).Value))=Trim$(customerName) Then
            oldID=CStr(ws.Cells(r,1).Value): ws.Cells(r,1)=newCustomerID
            WriteAuditLog "UPDATE","Customer_Master",newCustomerID,"CustomerID",oldID,newCustomerID,"Manager correction"
            Exit Sub
        End If
    Next r
    Err.Raise vbObjectError+1003,"IOMS",MsgText("CUSTOMER_NOT_FOUND")
End Sub
