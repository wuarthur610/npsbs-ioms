Attribute VB_Name = "modCard"
Option Explicit

Public Sub ValidateCardForCustomer(ByVal cardID As String, ByVal customerID As String, ByVal treatmentDate As Date)
    Dim ws As Worksheet, r As Long, pdate As Date, found As Boolean
    If Len(cardID)=0 Then Exit Sub
    Set ws=ThisWorkbook.Worksheets("Card_Master")
    For r=2 To LastDataRow(ws,1)
        If CStr(ws.Cells(r,1).Value)=cardID Then
            found=True
            If CStr(ws.Cells(r,2).Value)<>customerID Then Err.Raise vbObjectError+1251,"IOMS",MsgText("CARD_CUSTOMER")
            If IsDate(ws.Cells(r,3).Value) Then pdate=CDate(ws.Cells(r,3).Value) Else pdate=0
            If pdate>treatmentDate Then Err.Raise vbObjectError+1252,"IOMS",MsgText("CARD_DATE")
            Exit For
        End If
    Next r
    If Not found Then Err.Raise vbObjectError+1253,"IOMS",MsgText("NO_CARD")
End Sub

Public Function FindUsableCard(ByVal customerID As String, ByVal treatmentDate As Date) As String
    Dim ws As Worksheet, r As Long, lastRow As Long, bestDate As Date, cid As String, remain As Double, pdate As Date
    Set ws=ThisWorkbook.Worksheets("Card_Master"): lastRow=LastDataRow(ws,1)
    For r=2 To lastRow
        If CStr(ws.Cells(r,2).Value)=customerID And SafeNumber(ws.Cells(r,7).Value)>0 Then
            If IsDate(ws.Cells(r,3).Value) Then pdate=CDate(ws.Cells(r,3).Value) Else pdate=0
            If pdate<=treatmentDate Then
                If cid="" Or pdate>bestDate Then cid=CStr(ws.Cells(r,1).Value): bestDate=pdate
            End If
        End If
    Next r
    FindUsableCard=cid
End Function

Public Function DeductSessions(ByVal cardID As String, ByVal requested As Double, ByRef shortage As Double) As Boolean
    Dim ws As Worksheet, r As Long, remain As Double, used As Double, oldRemain As Double
    shortage=0: If Len(cardID)=0 Then Exit Function
    Set ws=ThisWorkbook.Worksheets("Card_Master")
    For r=2 To LastDataRow(ws,1)
        If CStr(ws.Cells(r,1).Value)=cardID Then
            oldRemain=SafeNumber(ws.Cells(r,7).Value): used=WorksheetFunction.Min(oldRemain,requested)
            ws.Cells(r,6).Value=SafeNumber(ws.Cells(r,6).Value)+used: ws.Cells(r,7).Value=oldRemain-used
            ValidateBalance SafeNumber(ws.Cells(r,7).Value): shortage=requested-used
            WriteAuditLog "UPDATE","Card_Master",cardID,"RemainingSessions",CStr(oldRemain),CStr(ws.Cells(r,7).Value),"Treatment deduction"
            DeductSessions=True: Exit Function
        End If
    Next r
End Function

Public Function HandleShortage(ByVal treatmentRecordID As String, ByVal customerID As String, ByVal customerName As String, ByVal shortage As Double, ByVal resolution As String, ByVal shortageAmount As Double, ByVal notes As String) As String
    Dim ws As Worksheet, r As Long, id As String
    ValidateShortageResolution shortage,resolution
    If shortage<=0 Then HandleShortage="None": Exit Function
    If resolution=TW_Charge() Then If shortageAmount<=0 Then Err.Raise vbObjectError+1201,"IOMS",MsgText("SHORTAGE_FEE")
    Set ws=EnsureSheet("Pending_Shortage",Array("PendingID","OriginalTreatmentRecordID","CustomerID","CustomerName","PendingSessions","CreatedDate","Status","SettledDate","SettledCardID","Notes","CreatedBy"))
    If resolution=TW_Pending() Then
        r=LastDataRow(ws,1)+1: id=NextID("PEND-",ws,1)
        ws.Cells(r,1)=id: ws.Cells(r,2)=treatmentRecordID: ws.Cells(r,3)=customerID: ws.Cells(r,4)=customerName: ws.Cells(r,5)=shortage: ws.Cells(r,6)=Date: ws.Cells(r,7)="Pending": ws.Cells(r,10)=notes: ws.Cells(r,11)=CurrentUser()
        WriteAuditLog "CREATE","Pending_Shortage",id,"PendingSessions","",CStr(shortage),"Pending next card"
        HandleShortage=id
    Else
        HandleShortage=resolution
    End If
End Function

Public Function GenerateCardID(ByVal prefix As String, ByVal purchaseDate As Date) As String
    Dim ws As Worksheet, r As Long, d As String, seq As Long, v As String
    Set ws=ThisWorkbook.Worksheets("Card_Master"): d=Format$(purchaseDate,"yymmdd")
    For r=2 To LastDataRow(ws,1)
        v=CStr(ws.Cells(r,1).Value)
        If Left$(v,Len(prefix)+6)=prefix & d Then seq=WorksheetFunction.Max(seq,Val(Right$(v,3)))
    Next r
    GenerateCardID=prefix & d & Format$(seq+1,"000")
End Function

' ============================================================
' PATCH-01 additions
' ============================================================

' Central CardType -> CardID prefix mapping.
' Confirmed by Arthur: Caimian=C, Huaxia=H, Hejia=W.
' Yingyou prefix is NOT yet confirmed -> deliberately raises an error
' instead of guessing. Do not hard-code a guess here.
Public Function CardTypeToPrefix(ByVal cardType As String) As String
    Dim t As String
    t = Trim$(cardType)
    If t = TW_CardCaimian() Then
        CardTypeToPrefix = "C"
    ElseIf t = TW_CardHuaxia() Then
        CardTypeToPrefix = "H"
    ElseIf t = TW_CardHejia() Then
        CardTypeToPrefix = "W"
    ElseIf t = TW_CardYingyou() Then
        Err.Raise vbObjectError + 1260, "IOMS", MsgText("CARD_PREFIX_UNCONFIRMED")
    Else
        Err.Raise vbObjectError + 1261, "IOMS", MsgText("CARD_TYPE_UNKNOWN")
    End If
End Function

Private Function CardIDExists(ByVal cardID As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Card_Master")
    For r = 2 To LastDataRow(ws, 1)
        If CStr(ws.Cells(r, 1).Value) = cardID Then CardIDExists = True: Exit Function
    Next r
End Function

Private Function CustomerExists(ByVal customerID As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Customer_Master")
    For r = 2 To LastDataRow(ws, 1)
        If CStr(ws.Cells(r, 1).Value) = customerID Then CustomerExists = True: Exit Function
    Next r
End Function

' Writes one new card into Card_Master.
' Column order (verified against the live workbook header row):
' 1 CardID | 2 CustomerID | 3 PurchaseDate | 4 PurchaseAmount | 5 OriginalSessions
' 6 UsedSessions | 7 RemainingSessions | 8 PaymentMethod | 9 CardStatus | 10 Notes
Public Function CreateCard(ByVal customerID As String, ByVal cardType As String, _
                            ByVal purchaseDate As Date, ByVal purchaseAmount As Double, _
                            ByVal originalSessions As Double, _
                            Optional ByVal paymentMethod As String = "", _
                            Optional ByVal cardID As String = "", _
                            Optional ByVal notes As String = "") As String
    Dim ws As Worksheet, r As Long, newID As String, prefix As String

    RequireValue customerID, MsgText("CUSTOMER_REQUIRED")
    If Not CustomerExists(customerID) Then Err.Raise vbObjectError + 1262, "IOMS", MsgText("CUSTOMER_NOT_FOUND")
    If originalSessions <= 0 Then Err.Raise vbObjectError + 1263, "IOMS", MsgText("SESSION_REQUIRED")
    If purchaseAmount < 0 Then Err.Raise vbObjectError + 1264, "IOMS", MsgText("AMOUNT_NEGATIVE")

    If Len(Trim$(cardID)) > 0 Then
        newID = Trim$(cardID)
    Else
        prefix = CardTypeToPrefix(cardType)
        newID = GenerateCardID(prefix, purchaseDate)
    End If
    If CardIDExists(newID) Then Err.Raise vbObjectError + 1265, "IOMS", MsgText("CARD_ID_DUPLICATE")

    Set ws = ThisWorkbook.Worksheets("Card_Master")
    r = LastDataRow(ws, 1) + 1
    ws.Cells(r, 1).Value = newID
    ws.Cells(r, 2).Value = customerID
    ws.Cells(r, 3).Value = purchaseDate
    ws.Cells(r, 4).Value = purchaseAmount
    ws.Cells(r, 5).Value = originalSessions
    ws.Cells(r, 6).Value = 0
    ws.Cells(r, 7).Value = originalSessions
    ws.Cells(r, 8).Value = paymentMethod
    ws.Cells(r, 9).Value = "Active"
    ws.Cells(r, 10).Value = notes

    WriteAuditLog "CREATE", "Card_Master", newID, "RemainingSessions", "", CStr(originalSessions), "New card purchase"
    CreateCard = newID
End Function