Attribute VB_Name = "modSecurity"
Option Explicit

Public Function HasPermission(ByVal roleName As String, ByVal actionName As String) As Boolean
    Select Case roleName
        Case "StoreManager": HasPermission=True
        Case "Therapist": HasPermission=(actionName="CreateTreatment" Or actionName="EditTreatmentNotes")
        Case Else: HasPermission=False
    End Select
End Function

Public Function GetCurrentUser() As String
    GetCurrentUser=CurrentUser()
End Function
