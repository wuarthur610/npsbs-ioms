Attribute VB_Name = "modValidation"
Option Explicit

Public Sub RequireValue(ByVal v As Variant, ByVal msg As String)
    If Len(Trim$(CStr(Nz(v, "")))) = 0 Then Err.Raise vbObjectError + 1100, "IOMS", msg
End Sub

Public Function NormalizeResolution(ByVal displayText As String) As String
    If displayText = TW_Charge() Then NormalizeResolution = "CHARGE": Exit Function
    If displayText = TW_Waive() Then NormalizeResolution = "WAIVE": Exit Function
    If displayText = TW_Pending() Then NormalizeResolution = "PENDING": Exit Function
    NormalizeResolution = UCase$(Trim$(displayText))
End Function

Public Function NormalizePayment(ByVal displayText As String) As String
    If displayText = TW_Cash() Then NormalizePayment = "CASH": Exit Function
    If displayText = TW_Bank() Then NormalizePayment = "BANK": Exit Function
    NormalizePayment = UCase$(Trim$(displayText))
End Function

Public Sub ValidateShortageResolution(ByVal shortage As Double, ByVal resolution As String)
    If shortage <= 0 Then Exit Sub
    Select Case NormalizeResolution(resolution)
        Case "CHARGE", "WAIVE", "PENDING"
        Case Else: Err.Raise vbObjectError + 1101, "IOMS", MsgText("SHORTAGE_RESOLUTION")
    End Select
End Sub

Public Sub ValidateSessionCount(ByVal sessions As Double)
    If sessions <= 0 Then Err.Raise vbObjectError + 1102, "IOMS", MsgText("SESSION_REQUIRED")
End Sub

Public Sub ValidatePaymentMethod(ByVal method As String)
    If Len(Trim$(method)) = 0 Then Exit Sub
    Select Case NormalizePayment(method)
        Case "CASH", "BANK"
        Case Else: Err.Raise vbObjectError + 1103, "IOMS", "Payment method invalid."
    End Select
End Sub

Public Sub ValidateTreatment(ByVal customerName As String, ByVal therapist As String, ByVal treatmentDate As Date, ByVal sessions As Double)
    RequireValue customerName, MsgText("CUSTOMER_REQUIRED")
    RequireValue therapist, MsgText("THERAPIST_REQUIRED")
    ValidateSessionCount sessions
    If treatmentDate = 0 Then Err.Raise vbObjectError + 1104, "IOMS", MsgText("DATE_REQUIRED")
End Sub

Public Sub ValidateBalance(ByVal balance As Double)
    If balance < 0 Then Err.Raise vbObjectError + 1105, "IOMS", MsgText("NO_NEGATIVE")
End Sub

Public Sub ValidateVarianceReason(ByVal varianceQty As Double, ByVal reason As String)
    If varianceQty <> 0 And Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 1106, "IOMS", MsgText("INV_REASON")
End Sub
