Attribute VB_Name = "modFormLabels"
Option Explicit

Public Sub ApplyTraditionalChineseLabels()
    On Error GoTo EH
    SetFormCaption "frmMainMenu", UText(&H5982&, &H5370&, &H690D&, &H79D8&, &H79D8&, &H4E09&, &H5CE1&, &H5317&, &H5927&, &H5E97&)
    SetFormCaption "frmTreatment", UText(&H6CBB&, &H7642&, &H8A18&, &H9304&)
    SetFormCaption "frmCustomer", UText(&H5BA2&, &H6236&, &H8CC7&, &H6599&)
    SetFormCaption "frmCard", UText(&H5957&, &H5361&, &H8CC7&, &H6599&)
    SetFormCaption "frmCardProduct", UText(&H5957&, &H5361&, &H7522&, &H54C1&, &H660E&, &H7D30&)
    SetFormCaption "frmShortage", UText(&H4E0D&, &H8DB3&, &H5802&, &H6578&, &H8655&, &H7406&)
    SetFormCaption "frmPendingShortage", UText(&H5F85&, &H88DC&, &H6263&, &H8655&, &H7406&)
    SetFormCaption "frmInventory", UText(&H76E4&, &H9EDE&, &H8A18&, &H9304&)
    RelabelTreatment
    RelabelMain
    MsgBox UText(&H8868&, &H55AE&, &H4E2D&, &H6587&, &H6A19&, &H7C64&, &H5DF2&, &H5B8C&, &H6210&), vbInformation, "NPSBS IOMS"
    Exit Sub
EH:
    MsgBox "Label update failed: " & Err.Description, vbExclamation, "NPSBS IOMS"
End Sub

Private Sub SetFormCaption(ByVal nm As String, ByVal cap As String)
    ThisWorkbook.VBProject.VBComponents(nm).Designer.Caption = cap
End Sub

Private Sub RelabelTreatment()
    Dim d As Object
    Set d = ThisWorkbook.VBProject.VBComponents("frmTreatment").Designer
    d.Controls("lblCustomerName").Caption = UText(&H5BA2&, &H6236&, &H59D3&, &H540D&)
    d.Controls("lblCustomerID").Caption = "CustomerID"
    d.Controls("lblDate").Caption = UText(&H8ABF&, &H7406&, &H65E5&, &H671F&)
    d.Controls("lblTherapist").Caption = UText(&H8ABF&, &H7406&, &H5E2B&)
    d.Controls("lblStart").Caption = UText(&H958B&, &H59CB&, &H6642&, &H9593&)
    d.Controls("lblEnd").Caption = UText(&H7D50&, &H675F&, &H6642&, &H9593&)
    d.Controls("lblSessions").Caption = UText(&H5802&, &H6578&)
    d.Controls("lblType").Caption = UText(&H7642&, &H7642&, &H985E&, &H578B&)
    d.Controls("lblBooking").Caption = UText(&H9810&, &H7D04&, &H985E&, &H578B&)
    d.Controls("lblCard").Caption = UText(&H5957&, &H5361&, &H865F&)
    d.Controls("lblPay").Caption = UText(&H4ED8&, &H6B3E&, &H65B9&, &H5F0F&)
    d.Controls("lblRevenue").Caption = UText(&H7642&, &H7A0B&, &H6536&, &H5165&)
    d.Controls("lblResolution").Caption = UText(&H4E0D&, &H8DB3&, &H5802&, &H6578&, &H8655&, &H7406&)
    d.Controls("lblShortageAmount").Caption = UText(&H4E0D&, &H8DB3&, &H6536&, &H8CBB&)
    d.Controls("lblNotes").Caption = UText(&H5099&, &H8A3B&)
    d.Controls("cmdSave").Caption = UText(&H5132&, &H5B58&)
    d.Controls("cmdCancel").Caption = UText(&H53D6&, &H6D88&)
End Sub

Private Sub RelabelMain()
    Dim d As Object
    Set d = ThisWorkbook.VBProject.VBComponents("frmMainMenu").Designer
    d.Controls("cmdTreatment").Caption = UText(&H65B0&, &H589E&, &H6CBB&, &H7642&, &H8A18&, &H9304&)
    d.Controls("cmdCustomer").Caption = UText(&H5BA2&, &H6236&, &H8CC7&, &H6599&)
    d.Controls("cmdCard").Caption = UText(&H5957&, &H5361&, &H7BA1&, &H7406&)
    d.Controls("cmdShortage").Caption = UText(&H4E0D&, &H8DB3&, &H5802&, &H6578&, &H8655&, &H7406&)
    d.Controls("cmdPending").Caption = UText(&H5F85&, &H88DC&, &H6263&, &H8655&, &H7406&)
    d.Controls("cmdInventory").Caption = UText(&H5EAB&, &H5B58&, &H76E4&, &H9EDE&)
    d.Controls("cmdCardProduct").Caption = UText(&H5957&, &H5361&, &H7522&, &H54C1&, &H660E&, &H7D30&)
End Sub
