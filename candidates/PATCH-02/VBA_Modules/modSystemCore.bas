Attribute VB_Name = "modSystemCore"
Option Explicit

Public Sub EnsureProductionSheets()
    EnsureSheet "Treatment_Record", Array("TreatmentRecordID", "TreatmentDate", "TreatmentStart_Taipei", "TreatmentEnd_Taipei", "CustomerName", "CustomerID", "Therapist", "TherapistID", "Sessions", "TreatmentType", "BookingType", "CardID", "CardUsageStatus", "ShortageResolution", "ShortageSessions", "TreatmentRevenue", "PaymentMethod", "CalendarUID", "Notes", "CreatedAt", "CreatedBy", "Status")
    EnsureSheet "Card_Usage_Detail", Array("UsageDetailID", "TreatmentRecordID", "CardID", "CustomerID", "UsageDate", "SessionsUsed", "UsageType", "ShortageSessions", "Notes", "CreatedAt", "CreatedBy", "Version", "BeforeSessions", "AfterSessions")
    EnsureSheet "Card_Product_Detail", Array("DetailID", "CardID", "VersionNo", "ProductCode", "ProductName", "Quantity", "UnitPrice", "LineAmount", "EffectiveFrom", "EffectiveTo", "Notes", "CreatedAt", "CreatedBy")
    EnsureSheet "Inventory", Array("InventoryRecordID", "CountDate", "ProductCode", "BookQty", "ActualQty", "VarianceQty", "VarianceReason", "Status", "Notes", "CreatedAt", "CreatedBy")
    EnsureSheet "Audit_Log", Array("AuditID", "AuditTime", "UserName", "ActionType", "TableName", "RecordID", "FieldName", "OldValue", "NewValue", "Reason")
    EnsureSheet "Pending_Shortage", Array("PendingID", "OriginalTreatmentRecordID", "CustomerID", "CustomerName", "PendingSessions", "CreatedDate", "Status", "SettledDate", "SettledCardID", "Notes", "CreatedBy")
    EnsureSheet "Sales_Record", Array("SalesID", "SaleDate", "CustomerID", "CustomerName", "SaleType", "ProductCode", "ProductName", "Quantity", "UnitPrice", "Amount", "PaymentMethod", "CardID", "Notes", "CreatedAt", "CreatedBy")
End Sub

Public Function EnsureSheet(ByVal sheetName As String, ByVal headers As Variant) As Worksheet
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    If Len(Trim$(CStr(ws.Cells(1,1).Value)))=0 Then
        For i=LBound(headers) To UBound(headers): ws.Cells(1,i+1).Value=headers(i): Next i
        ws.Rows(1).Font.Bold=True
    End If
    Set EnsureSheet = ws
End Function

' PATCH-02: safely append a header to one column of an ALREADY EXISTING sheet
' without touching any other column. EnsureSheet only writes headers when A1
' is empty, so it will not backfill a header for a newly appended column on a
' sheet that already has data/headers in columns to its left. This does.
Public Sub EnsureColumnHeader(ByVal ws As Worksheet, ByVal colIndex As Long, ByVal headerText As String)
    If Len(Trim$(CStr(ws.Cells(1, colIndex).Value))) = 0 Then
        ws.Cells(1, colIndex).Value = headerText
        ws.Cells(1, colIndex).Font.Bold = True
    End If
End Sub

Public Function NextID(ByVal prefix As String, ByVal ws As Worksheet, ByVal idCol As Long) As String
    Dim r As Long, lastRow As Long, maxN As Long, v As String, n As Long
    lastRow=ws.Cells(ws.Rows.Count,idCol).End(xlUp).Row
    For r=2 To lastRow
        v=CStr(ws.Cells(r,idCol).Value)
        If Left$(v,Len(prefix))=prefix Then
            n=Val(Mid$(v,Len(prefix)+1)): If n>maxN Then maxN=n
        End If
    Next r
    NextID=prefix & Format$(maxN+1,"000000")
End Function

Public Function LastDataRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastDataRow=ws.Cells(ws.Rows.Count,col).End(xlUp).Row
    If LastDataRow<1 Then LastDataRow=1
End Function

Public Function CurrentUser() As String
    CurrentUser=Environ$("Username")
    If Len(CurrentUser)=0 Then CurrentUser=Application.UserName
End Function

Public Function Nz(ByVal v As Variant, Optional ByVal fallback As Variant="") As Variant
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Nz=fallback Else Nz=v
End Function

Public Function SafeNumber(ByVal v As Variant) As Double
    If IsNumeric(v) Then SafeNumber=CDbl(v) Else SafeNumber=0
End Function

Public Function ParseDateTime(ByVal dateText As String, ByVal timeText As String) As Date
    If Not IsDate(dateText) Then Err.Raise vbObjectError+700,"IOMS","Invalid date."
    If Len(Trim$(timeText))=0 Then
        ParseDateTime=CDate(dateText)
    ElseIf IsDate(timeText) Then
        ParseDateTime=CDate(dateText)+TimeValue(timeText)
    Else
        Err.Raise vbObjectError+701,"IOMS","Invalid time."
    End If
End Function
