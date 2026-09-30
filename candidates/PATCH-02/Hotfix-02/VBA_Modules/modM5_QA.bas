Attribute VB_Name = "modM5_QA"
Option Explicit

Public Function QAStatus(ByVal expected As String, ByVal actual As String) As String
    If expected=actual Then QAStatus="PASS" Else QAStatus="FAIL"
End Function

Public Function ReleaseGateReady(ByVal failedTests As Long) As Boolean
    ReleaseGateReady=(failedTests=0)
End Function

Public Sub RunProductionSmokeTest()
    Dim failed As Long
    On Error GoTo EH
    EnsureProductionSheets
    If ThisWorkbook.Worksheets("Treatment_Record").Cells(1,1).Value<>"TreatmentRecordID" Then failed=failed+1
    If ThisWorkbook.Worksheets("Card_Usage_Detail").Cells(1,1).Value<>"UsageDetailID" Then failed=failed+1
    If ThisWorkbook.Worksheets("Audit_Log").Cells(1,1).Value<>"AuditID" Then failed=failed+1
    If failed=0 Then MsgBox MsgText("SMOKE_PASS"),vbInformation,"NPSBS IOMS V1.0" Else MsgBox MsgText("SMOKE_FAIL") & " / Failed=" & failed,vbCritical,"NPSBS IOMS V1.0"
    Exit Sub
EH:
    MsgBox MsgText("SMOKE_FAIL") & " - " & Err.Description,vbCritical,"NPSBS IOMS V1.0"
End Sub
