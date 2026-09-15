Attribute VB_Name = "modMainMenu"
Option Explicit

Public Sub IOMS_Initialize()
    EnsureProductionSheets
    MsgBox MsgText("INIT_OK"), vbInformation, "NPSBS IOMS V1.0"
End Sub

Public Sub BuildProductionUserForms()
    BuildAllProductionForms
End Sub

Private Sub ShowIOMSForm(ByVal formName As String)
    On Error GoTo EH
    VBA.UserForms.Add(formName).Show
    Exit Sub
EH:
    MsgBox "Unable to open form: " & formName & vbCrLf & _
           "Please run BuildProductionUserForms first." & vbCrLf & _
           "Error " & Err.Number & " / " & Err.Description, vbExclamation, "NPSBS IOMS"
End Sub

Public Sub OpenTreatmentForm()
    EnsureProductionSheets
    ShowIOMSForm "frmTreatment"
End Sub

Public Sub OpenCustomerForm()
    ShowIOMSForm "frmCustomer"
End Sub
Public Sub OpenCardForm()
    ShowIOMSForm "frmCard"
End Sub
Public Sub OpenShortageForm()
    ShowIOMSForm "frmShortage"
End Sub
Public Sub OpenPendingShortageForm()
    ShowIOMSForm "frmPendingShortage"
End Sub
Public Sub OpenCardProductForm()
    ShowIOMSForm "frmCardProduct"
End Sub
Public Sub OpenInventoryForm()
    ShowIOMSForm "frmInventory"
End Sub
Public Sub OpenMainMenu()
    EnsureProductionSheets
    ShowIOMSForm "frmMainMenu"
End Sub
