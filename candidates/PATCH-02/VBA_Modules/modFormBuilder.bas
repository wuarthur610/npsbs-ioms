Attribute VB_Name = "modFormBuilder"
Option Explicit

Private Const CT_MSFORM As Long = 3

Public Sub BuildAllProductionForms()
    Dim stage As String
    On Error GoTo EH
    If Not AccessVBProjectEnabled() Then
        MsgBox "Please enable Trust access to the VBA project object model, then run BuildProductionUserForms again.", vbCritical, "NPSBS IOMS"
        Exit Sub
    End If
    stage = "Remove old forms"
    RemoveFormIfExists "frmMainMenu"
    RemoveFormIfExists "frmTreatment"
    RemoveFormIfExists "frmCustomer"
    RemoveFormIfExists "frmCard"
    RemoveFormIfExists "frmCardProduct"
    RemoveFormIfExists "frmShortage"
    RemoveFormIfExists "frmPendingShortage"
    RemoveFormIfExists "frmInventory"
    ThisWorkbook.Save
    stage = "Build frmMainMenu": BuildMainMenu
    stage = "Build frmTreatment": BuildTreatment
    stage = "Build frmCustomer": BuildCustomer
    stage = "Build frmCard": BuildCard
    stage = "Build frmCardProduct": BuildCardProduct
    stage = "Build frmShortage": BuildShortage
    stage = "Build frmPendingShortage": BuildPendingShortage
    stage = "Build frmInventory": BuildInventory
    ThisWorkbook.Save
    MsgBox "Production UserForms created. They should now appear under Forms.", vbInformation, "NPSBS IOMS"
    Exit Sub
EH:
    MsgBox "Form build failed at: " & stage & vbCrLf & _
           "Error " & Err.Number & " / " & Err.Description, vbCritical, "NPSBS IOMS"
End Sub

Private Function AccessVBProjectEnabled() As Boolean
    Dim x As Object
    On Error Resume Next
    Set x = ThisWorkbook.VBProject.VBComponents
    AccessVBProjectEnabled = (Err.Number = 0 And Not x Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub RemoveFormIfExists(ByVal formName As String)
    Dim c As Object
    On Error Resume Next
    Set c = ThisWorkbook.VBProject.VBComponents(formName)
    If Not c Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove c
    Set c = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Function NewForm(ByVal formName As String, ByVal caption As String, ByVal w As Single, ByVal h As Single) As Object
    Dim c As Object
    Dim tmpName As String
    Set c = ThisWorkbook.VBProject.VBComponents.Add(CT_MSFORM)
    tmpName = GetTemporaryFormName()
    c.Name = tmpName
    SetComponentProperty c, "Caption", caption
    SetComponentProperty c, "Width", w
    SetComponentProperty c, "Height", h
    SetComponentProperty c, "StartUpPosition", 1
    ' Always start with a completely empty form code module.
    ' This prevents stale event procedures from surviving a rebuild.
    On Error Resume Next
    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines
    Err.Clear
    On Error GoTo 0
    ThisWorkbook.Save
    c.Name = formName
    ThisWorkbook.Save
    Set NewForm = c
End Function

Private Sub SetComponentProperty(ByVal component As Object, ByVal propertyName As String, ByVal value As Variant)
    On Error GoTo EH
    component.Properties(propertyName).Value = value
    Exit Sub
EH:
    Err.Raise Err.Number, "IOMS FormBuilder", "Cannot set UserForm property '" & propertyName & "': " & Err.Description
End Sub

Private Function GetTemporaryFormName() As String
    Dim n As Long
    Dim candidate As String
    n = 1
    Do
        candidate = "frmIOMS_BuildTmp" & CStr(n)
        If Not ComponentExists(candidate) Then
            GetTemporaryFormName = candidate
            Exit Function
        End If
        n = n + 1
    Loop
End Function

Private Function ComponentExists(ByVal componentName As String) As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = ThisWorkbook.VBProject.VBComponents(componentName)
    ComponentExists = Not c Is Nothing
    Set c = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function AddCtl(ByVal c As Object, ByVal progId As String, ByVal nm As String, ByVal left As Single, ByVal top As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim ctl As Object
    On Error GoTo EH
    Set ctl = c.Designer.Controls.Add(progId, nm, True)
    ctl.Left = left
    ctl.Top = top
    ctl.Width = w
    ctl.Height = h
    Set AddCtl = ctl
    Exit Function
EH:
    Err.Raise Err.Number, "IOMS FormBuilder", "Cannot add control '" & nm & "': " & Err.Description
End Function

Private Sub AddLabel(ByVal c As Object, ByVal nm As String, ByVal cap As String, ByVal top As Single)
    Dim x As Object
    Set x = AddCtl(c, "Forms.Label.1", nm, 18, top, 100, 18)
    x.Caption = cap
End Sub

Private Sub AddText(ByVal c As Object, ByVal nm As String, ByVal top As Single, Optional ByVal w As Single = 190)
    Call AddCtl(c, "Forms.TextBox.1", nm, 125, top, w, 20)
End Sub

Private Sub AddCombo(ByVal c As Object, ByVal nm As String, ByVal top As Single, Optional ByVal w As Single = 190)
    Call AddCtl(c, "Forms.ComboBox.1", nm, 125, top, w, 20)
End Sub

Private Sub AddButton(ByVal c As Object, ByVal nm As String, ByVal cap As String, ByVal left As Single, ByVal top As Single, ByVal w As Single, ByVal handler As String)
    Dim x As Object
    Dim cm As Object
    Dim code As String
    Set x = AddCtl(c, "Forms.CommandButton.1", nm, left, top, w, 26)
    x.Caption = cap
    Set cm = c.CodeModule
    ' Do not add a duplicate event procedure if the form already contains one.
    Dim procLine As Long
    Dim procLines As Long
    On Error Resume Next
    Do
        procLine = cm.ProcStartLine(nm & "_Click", 0)
        procLines = cm.ProcCountLines(nm & "_Click", 0)
        If procLine > 0 And procLines > 0 Then
            cm.DeleteLines procLine, procLines
        Else
            Exit Do
        End If
    Loop
    Err.Clear
    On Error GoTo 0
    ' Only 2 continuations here - well inside the VBA line-continuation limit,
    ' left as-is (Hotfix-01 scope: fix statements that actually exceed the
    ' limit, not rewrite statements that are already safe).
    code = "Private Sub " & nm & "_Click()" & vbCrLf & _
           "    " & handler & vbCrLf & _
           "End Sub" & vbCrLf
    cm.AddFromString code
End Sub

Private Function DQ() As String
    DQ = Chr$(34)
End Function

' ============================================================
' PATCH-02 Hotfix-01: line-continuation safe code assembly.
' VBA/VBE has a hard limit of 24 "_" continuations per logical statement.
' BuildTreatment (24) was already at that limit and BuildShortage (35)
' exceeded it, which is exactly what produced the "line-continuation
' character used too many times" error on import.
' AppendCode replaces every long "code = ... & vbCrLf & _" chain with a
' sequence of independent single-line statements, each with ZERO
' continuations, so this class of error cannot recur no matter how many
' lines a generated form needs. The longest single generated line (the
' CreateTreatmentRecord call in BuildTreatment) is ~400 characters as an
' AppendCode call site, well under VBA's separate 1023-character line
' length limit.
' ============================================================
Private Sub AppendCode(ByRef code As String, ByVal lineText As String)
    code = code & lineText & vbCrLf
End Sub

Private Sub BuildMainMenu()
    Dim c As Object, y As Single
    Set c = NewForm("frmMainMenu", "NPSBS IOMS V1.0", 330, 330)
    y = 24
    AddButton c, "cmdTreatment", "Open Treatment", 30, y, 250, "OpenTreatmentForm": y = y + 38
    AddButton c, "cmdCustomer", "Open Customer", 30, y, 250, "OpenCustomerForm": y = y + 38
    AddButton c, "cmdCard", "Open Card", 30, y, 250, "OpenCardForm": y = y + 38
    AddButton c, "cmdShortage", "Open Shortage", 30, y, 250, "OpenShortageForm": y = y + 38
    AddButton c, "cmdPending", "Open Pending Settlement", 30, y, 250, "OpenPendingShortageForm": y = y + 38
    AddButton c, "cmdInventory", "Open Inventory", 30, y, 250, "OpenInventoryForm": y = y + 38
    AddButton c, "cmdCardProduct", "Open Card Product", 30, y, 250, "OpenCardProductForm"
End Sub

Private Sub BuildTreatment()
    Dim c As Object, y As Single, cm As Object, code As String
    Set c = NewForm("frmTreatment", "Treatment Entry", 470, 570)
    y = 20
    AddLabel c, "lblCustomerName", "Customer", y: AddText c, "txtCustomerName", y: y = y + 32
    AddLabel c, "lblCustomerID", "Customer ID", y: AddText c, "txtCustomerID", y: y = y + 32
    AddLabel c, "lblDate", "Date", y: AddText c, "txtTreatmentDate", y: y = y + 32
    AddLabel c, "lblTherapist", "Therapist", y: AddCombo c, "cboTherapist", y: y = y + 32
    AddLabel c, "lblStart", "Start", y: AddText c, "txtStart", y: y = y + 32
    AddLabel c, "lblEnd", "End", y: AddText c, "txtEnd", y: y = y + 32
    AddLabel c, "lblSessions", "Sessions", y: AddText c, "txtSessions", y: y = y + 32
    AddLabel c, "lblType", "Treatment Type", y: AddText c, "txtTreatmentType", y: y = y + 32
    AddLabel c, "lblBooking", "Booking Type", y: AddText c, "txtBookingType", y: y = y + 32
    AddLabel c, "lblCard", "Card ID", y: AddCombo c, "cboCardID", y: y = y + 32
    AddLabel c, "lblPay", "Payment", y: AddCombo c, "cboPayment", y: y = y + 32
    AddLabel c, "lblRevenue", "Revenue", y: AddText c, "txtRevenue", y: y = y + 32
    AddLabel c, "lblResolution", "Shortage", y: AddCombo c, "cboShortage", y: y = y + 32
    AddLabel c, "lblShortageAmount", "Shortage Fee", y: AddText c, "txtShortageAmount", y: y = y + 32
    AddLabel c, "lblNotes", "Notes", y: AddText c, "txtNotes", y, 250: y = y + 40
    AddButton c, "cmdSave", "Save", 125, y, 110, "SaveTreatmentForm"
    AddButton c, "cmdCancel", "Cancel", 245, y, 110, "Unload Me"
    Set cm = c.CodeModule
    code = ""
    AppendCode code, "Private Sub UserForm_Initialize()"
    AppendCode code, "    Me.txtTreatmentDate.Value = Format$(Date, " & DQ() & "yyyy-mm-dd" & DQ() & ")"
    AppendCode code, "    Me.cboPayment.AddItem TW_Cash()"
    AppendCode code, "    Me.cboPayment.AddItem TW_Bank()"
    AppendCode code, "    Me.cboShortage.AddItem " & DQ() & DQ()
    AppendCode code, "    Me.cboShortage.AddItem TW_Charge()"
    AppendCode code, "    Me.cboShortage.AddItem TW_Waive()"
    AppendCode code, "    Me.cboShortage.AddItem TW_Pending()"
    AppendCode code, "End Sub"
    AppendCode code, "Private Sub txtCustomerName_Exit(ByVal Cancel As MSForms.ReturnBoolean)"
    AppendCode code, "    On Error Resume Next"
    AppendCode code, "    Me.txtCustomerID.Value = GetCustomerByName(Me.txtCustomerName.Value)"
    AppendCode code, "End Sub"
    AppendCode code, "Public Sub SaveTreatmentForm()"
    AppendCode code, "    Dim rid As String, res As String, pay As String"
    AppendCode code, "    On Error GoTo EH"
    AppendCode code, "    If Me.cboPayment.ListIndex = 0 Then pay = TW_Cash() Else pay = TW_Bank()"
    AppendCode code, "    If Me.cboShortage.ListIndex > 0 Then res = Me.cboShortage.Value Else res = " & DQ() & DQ()
    AppendCode code, "    rid = CreateTreatmentRecord(CDate(Me.txtTreatmentDate.Value), TimeValue(Me.txtStart.Value), TimeValue(Me.txtEnd.Value), Me.txtCustomerName.Value, Me.cboTherapist.Value, " & DQ() & DQ() & ", CDbl(Me.txtSessions.Value), Me.txtTreatmentType.Value, Me.txtBookingType.Value, Me.cboCardID.Value, pay, CDbl(Val(Me.txtRevenue.Value)), " & DQ() & DQ() & ", Me.txtNotes.Value, res, CDbl(Val(Me.txtShortageAmount.Value)))"
    AppendCode code, "    MsgBox " & DQ() & "Saved: " & DQ() & " & rid, vbInformation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "    Unload Me"
    AppendCode code, "    Exit Sub"
    AppendCode code, "EH:"
    AppendCode code, "    MsgBox Err.Description, vbExclamation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "End Sub"
    cm.AddFromString code
End Sub

' ============================================================
' PATCH-01: Runtime binding for the six remaining UserForms.
' Module-level placeholder Save* subs removed; each form now owns
' its Public Sub Save*Form in its own code module.
' PATCH-02 Hotfix-01: code assembly converted from long "& vbCrLf & _"
' chains to AppendCode (see note above BuildMainMenu).
' ============================================================

Private Sub BuildCustomer()
    Dim c As Object, y As Single, cm As Object, code As String
    Set c = NewForm("frmCustomer", "Customer", 430, 260): y = 24
    AddLabel c, "lblName", "Customer", y: AddText c, "txtName", y: y = y + 34
    AddLabel c, "lblNotes", "Notes", y: AddText c, "txtNotes", y: y = y + 44
    AddButton c, "cmdSave", "Save", 125, y, 110, "SaveCustomerForm"
    AddButton c, "cmdCancel", "Cancel", 245, y, 110, "Unload Me"
    Set cm = c.CodeModule
    code = ""
    AppendCode code, "Public Sub SaveCustomerForm()"
    AppendCode code, "    Dim newID As String"
    AppendCode code, "    On Error GoTo EH"
    AppendCode code, "    newID = CreateCustomer(Me.txtName.Value, Me.txtNotes.Value)"
    AppendCode code, "    MsgBox " & DQ() & "Saved: " & DQ() & " & newID, vbInformation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "    Unload Me"
    AppendCode code, "    Exit Sub"
    AppendCode code, "EH:"
    AppendCode code, "    MsgBox Err.Description, vbExclamation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "End Sub"
    cm.AddFromString code
End Sub

Private Sub BuildCard()
    Dim c As Object, y As Single, cm As Object, code As String
    Set c = NewForm("frmCard", "Card", 450, 360): y = 24
    AddLabel c, "lblCardID", "Card ID (optional)", y: AddText c, "txtCardID", y: y = y + 32
    AddLabel c, "lblCustomerID", "Customer ID", y: AddText c, "txtCustomerID", y: y = y + 32
    AddLabel c, "lblPurchase", "Purchase Date", y: AddText c, "txtPurchaseDate", y: y = y + 32
    AddLabel c, "lblAmount", "Amount", y: AddText c, "txtAmount", y: y = y + 32
    AddLabel c, "lblSessions", "Sessions", y: AddText c, "txtSessions", y: y = y + 32
    AddLabel c, "lblCardType", "Card Type", y: AddCombo c, "cboCardType", y: y = y + 32
    AddLabel c, "lblPayment", "Payment", y: AddCombo c, "cboPayment", y: y = y + 32
    AddButton c, "cmdSave", "Save", 125, y, 110, "SaveCardForm"
    AddButton c, "cmdCancel", "Cancel", 245, y, 110, "Unload Me"
    Set cm = c.CodeModule
    code = ""
    AppendCode code, "Private Sub UserForm_Initialize()"
    AppendCode code, "    Me.txtPurchaseDate.Value = Format$(Date, " & DQ() & "yyyy-mm-dd" & DQ() & ")"
    AppendCode code, "    Me.cboCardType.AddItem TW_CardCaimian()"
    AppendCode code, "    Me.cboCardType.AddItem TW_CardHuaxia()"
    AppendCode code, "    Me.cboCardType.AddItem TW_CardHejia()"
    AppendCode code, "    Me.cboPayment.AddItem TW_Cash()"
    AppendCode code, "    Me.cboPayment.AddItem TW_Bank()"
    AppendCode code, "End Sub"
    AppendCode code, "Public Sub SaveCardForm()"
    AppendCode code, "    Dim newID As String, pay As String"
    AppendCode code, "    On Error GoTo EH"
    AppendCode code, "    If Me.cboPayment.ListIndex = 0 Then pay = TW_Cash() Else pay = TW_Bank()"
    AppendCode code, "    If Len(Trim$(Me.txtCardID.Value)) = 0 And Me.cboCardType.ListIndex < 0 Then Err.Raise vbObjectError + 1266, " & DQ() & "IOMS" & DQ() & ", MsgText(" & DQ() & "CARD_TYPE_REQUIRED" & DQ() & ")"
    AppendCode code, "    newID = CreateCard(Me.txtCustomerID.Value, CStr(Me.cboCardType.Value), CDate(Me.txtPurchaseDate.Value), CDbl(Val(Me.txtAmount.Value)), CDbl(Val(Me.txtSessions.Value)), pay, Me.txtCardID.Value, " & DQ() & DQ() & ")"
    AppendCode code, "    MsgBox " & DQ() & "Saved: " & DQ() & " & newID, vbInformation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "    Unload Me"
    AppendCode code, "    Exit Sub"
    AppendCode code, "EH:"
    AppendCode code, "    MsgBox Err.Description, vbExclamation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "End Sub"
    cm.AddFromString code
End Sub

Private Sub BuildCardProduct()
    Dim c As Object, y As Single, cm As Object, code As String
    Set c = NewForm("frmCardProduct", "Card Product", 470, 300): y = 24
    AddLabel c, "lblCardID", "Card ID", y: AddText c, "txtCardID", y: y = y + 32
    AddLabel c, "lblProduct", "Product Code", y: AddText c, "txtProductCode", y: y = y + 32
    AddLabel c, "lblName", "Product Name", y: AddText c, "txtProductName", y: y = y + 32
    AddLabel c, "lblQty", "Quantity", y: AddText c, "txtQuantity", y: y = y + 32
    AddLabel c, "lblPrice", "Unit Price", y: AddText c, "txtUnitPrice", y: y = y + 32
    AddButton c, "cmdSave", "Save", 125, y, 110, "SaveCardProductForm"
    AddButton c, "cmdCancel", "Cancel", 245, y, 110, "Unload Me"
    Set cm = c.CodeModule
    code = ""
    AppendCode code, "Public Sub SaveCardProductForm()"
    AppendCode code, "    Dim newID As String"
    AppendCode code, "    On Error GoTo EH"
    AppendCode code, "    newID = SaveCardProductVersion(Me.txtCardID.Value, Me.txtProductCode.Value, Me.txtProductName.Value, CDbl(Val(Me.txtQuantity.Value)), CDbl(Val(Me.txtUnitPrice.Value)), " & DQ() & DQ() & ")"
    AppendCode code, "    MsgBox " & DQ() & "Saved: " & DQ() & " & newID, vbInformation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "    Unload Me"
    AppendCode code, "    Exit Sub"
    AppendCode code, "EH:"
    AppendCode code, "    MsgBox Err.Description, vbExclamation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "End Sub"
    cm.AddFromString code
End Sub

Private Sub BuildShortage()
    Dim c As Object, y As Single, cm As Object, code As String
    Set c = NewForm("frmShortage", "Shortage Resolution", 450, 320): y = 24
    AddLabel c, "lblTreatment", "Treatment ID", y: AddText c, "txtTreatmentID", y: y = y + 34
    AddLabel c, "lblSessions", "Shortage Sessions", y: AddText c, "txtSessions", y: y = y + 34
    AddLabel c, "lblResolution", "Resolution", y: AddCombo c, "cboResolution", y: y = y + 34
    AddLabel c, "lblShortageAmount", "Shortage Fee", y: AddText c, "txtShortageAmount", y: y = y + 44
    AddButton c, "cmdSave", "Save", 125, y, 110, "SaveShortageForm"
    AddButton c, "cmdCancel", "Cancel", 245, y, 110, "Unload Me"
    Set cm = c.CodeModule
    code = ""
    AppendCode code, "Private Sub UserForm_Initialize()"
    AppendCode code, "    Me.cboResolution.AddItem TW_Charge()"
    AppendCode code, "    Me.cboResolution.AddItem TW_Waive()"
    AppendCode code, "    Me.cboResolution.AddItem TW_Pending()"
    AppendCode code, "    Me.txtSessions.Locked = True"
    AppendCode code, "End Sub"
    AppendCode code, "Private Sub txtTreatmentID_Exit(ByVal Cancel As MSForms.ReturnBoolean)"
    AppendCode code, "    On Error Resume Next"
    AppendCode code, "    Dim ws As Worksheet, r As Long"
    AppendCode code, "    Set ws = ThisWorkbook.Worksheets(" & DQ() & "Treatment_Record" & DQ() & ")"
    AppendCode code, "    For r = 2 To LastDataRow(ws, 1)"
    AppendCode code, "        If CStr(ws.Cells(r, 1).Value) = Me.txtTreatmentID.Value Then"
    AppendCode code, "            Me.txtSessions.Value = CStr(ws.Cells(r, 15).Value): Exit For"
    AppendCode code, "        End If"
    AppendCode code, "    Next r"
    AppendCode code, "End Sub"
    AppendCode code, "Public Sub SaveShortageForm()"
    AppendCode code, "    Dim ws As Worksheet, r As Long, custID As String, custName As String, shortageQty As Double, pendingID As String"
    AppendCode code, "    On Error GoTo EH"
    AppendCode code, "    Set ws = ThisWorkbook.Worksheets(" & DQ() & "Treatment_Record" & DQ() & ")"
    AppendCode code, "    For r = 2 To LastDataRow(ws, 1)"
    AppendCode code, "        If CStr(ws.Cells(r, 1).Value) = Me.txtTreatmentID.Value Then"
    AppendCode code, "            custID = CStr(ws.Cells(r, 6).Value): custName = CStr(ws.Cells(r, 5).Value): shortageQty = SafeNumber(ws.Cells(r, 15).Value): Exit For"
    AppendCode code, "        End If"
    AppendCode code, "    Next r"
    AppendCode code, "    If Len(custID) = 0 Then Err.Raise vbObjectError + 1601, " & DQ() & "IOMS" & DQ() & ", " & DQ() & "Treatment record not found." & DQ()
    AppendCode code, "    If Len(Trim$(CStr(ws.Cells(r, 14).Value))) > 0 Then Err.Raise vbObjectError + 1602, " & DQ() & "IOMS" & DQ() & ", MsgText(" & DQ() & "SHORTAGE_ALREADY_RESOLVED" & DQ() & ")"
    AppendCode code, "    If shortageQty <= 0 Then Err.Raise vbObjectError + 1603, " & DQ() & "IOMS" & DQ() & ", MsgText(" & DQ() & "NO_SHORTAGE" & DQ() & ")"
    AppendCode code, "    pendingID = HandleShortage(Me.txtTreatmentID.Value, custID, custName, shortageQty, Me.cboResolution.Value, CDbl(Val(Me.txtShortageAmount.Value)), " & DQ() & DQ() & ")"
    AppendCode code, "    ws.Cells(r, 14).Value = Me.cboResolution.Value"
    AppendCode code, "    MsgBox " & DQ() & "Resolved: " & DQ() & " & pendingID, vbInformation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "    Unload Me"
    AppendCode code, "    Exit Sub"
    AppendCode code, "EH:"
    AppendCode code, "    MsgBox Err.Description, vbExclamation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "End Sub"
    cm.AddFromString code
End Sub

Private Sub BuildPendingShortage()
    Dim c As Object, y As Single, cm As Object, code As String
    Set c = NewForm("frmPendingShortage", "Pending Settlement", 450, 260): y = 24
    AddLabel c, "lblPending", "Pending ID", y: AddText c, "txtPendingID", y: y = y + 34
    AddLabel c, "lblCard", "New Card ID", y: AddText c, "txtNewCardID", y: y = y + 34
    AddLabel c, "lblSessions", "Sessions", y: AddText c, "txtSessions", y: y = y + 44
    AddButton c, "cmdSave", "Settle", 125, y, 110, "SavePendingForm"
    AddButton c, "cmdCancel", "Cancel", 245, y, 110, "Unload Me"
    Set cm = c.CodeModule
    code = ""
    AppendCode code, "Public Sub SavePendingForm()"
    AppendCode code, "    On Error GoTo EH"
    AppendCode code, "    SettlePending Me.txtPendingID.Value, Me.txtNewCardID.Value, CDbl(Val(Me.txtSessions.Value))"
    AppendCode code, "    MsgBox " & DQ() & "Settled." & DQ() & ", vbInformation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "    Unload Me"
    AppendCode code, "    Exit Sub"
    AppendCode code, "EH:"
    AppendCode code, "    MsgBox Err.Description, vbExclamation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "End Sub"
    cm.AddFromString code
End Sub

Private Sub BuildInventory()
    Dim c As Object, y As Single, cm As Object, code As String
    Set c = NewForm("frmInventory", "Inventory", 450, 300): y = 24
    AddLabel c, "lblProduct", "Product Code", y: AddText c, "txtProductCode", y: y = y + 34
    AddLabel c, "lblDate", "Count Date", y: AddText c, "txtCountDate", y: y = y + 34
    AddLabel c, "lblActual", "Actual Qty", y: AddText c, "txtActualQty", y: y = y + 34
    AddLabel c, "lblReason", "Variance Reason", y: AddText c, "txtReason", y: y = y + 34
    AddButton c, "cmdSave", "Save", 125, y, 110, "SaveInventoryForm"
    AddButton c, "cmdCancel", "Cancel", 245, y, 110, "Unload Me"
    Set cm = c.CodeModule
    code = ""
    AppendCode code, "Private Sub UserForm_Initialize()"
    AppendCode code, "    Me.txtCountDate.Value = Format$(Date, " & DQ() & "yyyy-mm-dd" & DQ() & ")"
    AppendCode code, "End Sub"
    AppendCode code, "Public Sub SaveInventoryForm()"
    AppendCode code, "    Dim newID As String"
    AppendCode code, "    On Error GoTo EH"
    AppendCode code, "    newID = ProcessCountVariance(Me.txtProductCode.Value, CDate(Me.txtCountDate.Value), CDbl(Val(Me.txtActualQty.Value)), Me.txtReason.Value, " & DQ() & DQ() & ")"
    AppendCode code, "    MsgBox " & DQ() & "Saved: " & DQ() & " & newID, vbInformation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "    Unload Me"
    AppendCode code, "    Exit Sub"
    AppendCode code, "EH:"
    AppendCode code, "    MsgBox Err.Description, vbExclamation, " & DQ() & "NPSBS IOMS" & DQ()
    AppendCode code, "End Sub"
    cm.AddFromString code
End Sub
