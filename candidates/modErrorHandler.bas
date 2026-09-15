Attribute VB_Name = "modErrorHandler"
Option Explicit

Public Sub HandleError(ByVal procName As String, ByVal errNumber As Long, ByVal errDescription As String)
    WriteAuditLog "ERROR","SYSTEM",procName,"Error",CStr(errNumber),errDescription
    MsgBox "IOMS Error: " & vbCrLf & errDescription & vbCrLf & "Procedure: " & procName,vbExclamation,"NPSBS IOMS V1.0"
End Sub

Public Sub RollbackTransaction()
    'Single-row operations are committed atomically at the record level; this hook is reserved for future multi-row rollback.
End Sub
