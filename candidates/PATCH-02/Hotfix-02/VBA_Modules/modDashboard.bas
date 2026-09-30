Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub RefreshIOMSDashboard()
    On Error Resume Next
    ThisWorkbook.Worksheets("M5_10_Dashboard_Spec").Calculate
    ThisWorkbook.Worksheets("M5_10_Therapist_KPI").Calculate
    On Error GoTo 0
    MsgBox "Dashboard / KPI refreshed.",vbInformation,"NPSBS IOMS V1.0"
End Sub
