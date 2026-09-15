Attribute VB_Name = "modAudit"
Option Explicit

Public Sub WriteAuditLog(ByVal actionType As String, ByVal tableName As String, ByVal recordID As String, ByVal fieldName As String, ByVal oldValue As String, ByVal newValue As String, Optional ByVal reason As String="")
    Dim ws As Worksheet, r As Long
    Set ws=EnsureSheet("Audit_Log",Array("AuditID","AuditTime","UserName","ActionType","TableName","RecordID","FieldName","OldValue","NewValue","Reason"))
    r=LastDataRow(ws,1)+1
    ws.Cells(r,1).Value="AUD-" & Format$(Now,"yyyymmddhhmmss") & "-" & Format$(r,"0000")
    ws.Cells(r,2).Value=Now: ws.Cells(r,3).Value=CurrentUser(): ws.Cells(r,4).Value=actionType
    ws.Cells(r,5).Value=tableName: ws.Cells(r,6).Value=recordID: ws.Cells(r,7).Value=fieldName
    ws.Cells(r,8).Value=oldValue: ws.Cells(r,9).Value=newValue: ws.Cells(r,10).Value=reason
End Sub

Public Function GetAuditHistory(ByVal tableName As String, ByVal recordID As String) As String
    Dim ws As Worksheet, r As Long, lastRow As Long, s As String
    Set ws=EnsureSheet("Audit_Log",Array("AuditID","AuditTime","UserName","ActionType","TableName","RecordID","FieldName","OldValue","NewValue","Reason"))
    lastRow=LastDataRow(ws,1)
    For r=2 To lastRow
        If CStr(ws.Cells(r,5).Value)=tableName And CStr(ws.Cells(r,6).Value)=recordID Then s=s & ws.Cells(r,2).Text & " | " & ws.Cells(r,4).Value & " | " & ws.Cells(r,7).Value & vbCrLf
    Next r
    GetAuditHistory=s
End Function
