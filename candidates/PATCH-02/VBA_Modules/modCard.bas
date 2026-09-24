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

' PATCH-02 / P1-01: select the OLDEST eligible card first (was: latest).
' Eligibility unchanged: CustomerID matches, PurchaseDate <= TreatmentDate,
' RemainingSessions > 0. (CardStatus is not yet a physical column on the live
' sheet at time of writing; see Implementation Report "Known limitations".)
Public Function FindUsableCard(ByVal customerID As String, ByVal treatmentDate As Date) As String
    Dim ws As Worksheet, r As Long, lastRow As Long, bestDate As Date, cid As String, pdate As Date, found As Boolean
    Set ws=ThisWorkbook.Worksheets("Card_Master"): lastRow=LastDataRow(ws,1)
    For r=2 To lastRow
        If CStr(ws.Cells(r,2).Value)=customerID And SafeNumber(ws.Cells(r,7).Value)>0 Then
            If IsDate(ws.Cells(r,3).Value) Then pdate=CDate(ws.Cells(r,3).Value) Else pdate=0
            If pdate<=treatmentDate Then
                If Not found Or pdate<bestDate Then
                    cid=CStr(ws.Cells(r,1).Value): bestDate=pdate: found=True
                End If
            End If
        End If
    Next r
    FindUsableCard=cid
End Function

' ============================================================
' PATCH-02 / P0-02, P0-03: snapshot-based deduction primitives.
' CalculateDeduction is a PURE read - it performs no mutation, so callers can
' run it during the validation phase (before the first persistent write) and
' use its output both for the Balance Guard decision and as the transaction
' snapshot for later compensation if something downstream fails.
' CommitDeduction is the only place that actually writes Card_Master for a
' deduction; it re-checks the guard immediately before writing (defense in
' depth against a stale snapshot).
' ============================================================

' Pure lookup: current UsedSessions/RemainingSessions for a card. No mutation.
Public Function GetCardBalance(ByVal cardID As String, ByRef usedSessions As Double, ByRef remainingSessions As Double) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws=ThisWorkbook.Worksheets("Card_Master")
    For r=2 To LastDataRow(ws,1)
        If CStr(ws.Cells(r,1).Value)=cardID Then
            usedSessions=SafeNumber(ws.Cells(r,6).Value)
            remainingSessions=SafeNumber(ws.Cells(r,7).Value)
            GetCardBalance=True
            Exit Function
        End If
    Next r
End Function

' Pure calculation of what a deduction WOULD produce. No mutation.
' This is the single place the Balance Guard rule is enforced:
'   usedSessions   <= beforeRemaining
'   afterRemaining >= 0
Public Function CalculateDeduction(ByVal cardID As String, ByVal requested As Double, _
        ByRef beforeUsedTotal As Double, ByRef beforeRemaining As Double, _
        ByRef usedSessions As Double, ByRef shortageSessions As Double, _
        ByRef afterRemaining As Double) As Boolean
    beforeUsedTotal=0: beforeRemaining=0: usedSessions=0: shortageSessions=requested: afterRemaining=0
    If Len(cardID)=0 Then Exit Function
    If Not GetCardBalance(cardID, beforeUsedTotal, beforeRemaining) Then Exit Function
    usedSessions=WorksheetFunction.Min(beforeRemaining, requested)
    shortageSessions=requested-usedSessions
    afterRemaining=beforeRemaining-usedSessions
    If usedSessions>beforeRemaining Then Err.Raise vbObjectError+1270,"IOMS",MsgText("BALANCE_GUARD")
    If afterRemaining<0 Then Err.Raise vbObjectError+1270,"IOMS",MsgText("BALANCE_GUARD")
    CalculateDeduction=True
End Function

' Commits a deduction previously computed by CalculateDeduction IN THE SAME
' transaction (same snapshot values). Re-validates the guard before writing.
Public Sub CommitDeduction(ByVal cardID As String, ByVal beforeRemaining As Double, _
        ByVal usedSessions As Double, ByVal afterRemaining As Double)
    Dim ws As Worksheet, r As Long
    If usedSessions>beforeRemaining Then Err.Raise vbObjectError+1270,"IOMS",MsgText("BALANCE_GUARD")
    If afterRemaining<0 Then Err.Raise vbObjectError+1270,"IOMS",MsgText("BALANCE_GUARD")
    Set ws=ThisWorkbook.Worksheets("Card_Master")
    EnsureCardMasterAuditColumns ws
    For r=2 To LastDataRow(ws,1)
        If CStr(ws.Cells(r,1).Value)=cardID Then
            ws.Cells(r,6).Value=SafeNumber(ws.Cells(r,6).Value)+usedSessions
            ws.Cells(r,7).Value=afterRemaining
            ws.Cells(r,14).Value=Now
            ws.Cells(r,15).Value=CurrentUser()
            WriteAuditLog "UPDATE","Card_Master",cardID,"RemainingSessions",CStr(beforeRemaining),CStr(afterRemaining),"Treatment deduction"
            Exit Sub
        End If
    Next r
    Err.Raise vbObjectError+1271,"IOMS",MsgText("NO_CARD")
End Sub

' Compensating transaction: restores Card_Master to a previously captured
' snapshot (beforeUsedTotal/beforeRemaining from CalculateDeduction). Used
' when a later step in the same transaction fails after CommitDeduction
' already ran. Never deletes anything; always leaves an Audit_Log trail.
Public Sub RestoreCardSnapshot(ByVal cardID As String, ByVal snapshotUsedTotal As Double, _
        ByVal snapshotRemaining As Double, ByVal reason As String)
    Dim ws As Worksheet, r As Long, curRemain As Double
    If Len(cardID)=0 Then Exit Sub
    Set ws=ThisWorkbook.Worksheets("Card_Master")
    EnsureCardMasterAuditColumns ws
    For r=2 To LastDataRow(ws,1)
        If CStr(ws.Cells(r,1).Value)=cardID Then
            curRemain=SafeNumber(ws.Cells(r,7).Value)
            ws.Cells(r,6).Value=snapshotUsedTotal
            ws.Cells(r,7).Value=snapshotRemaining
            ws.Cells(r,14).Value=Now
            ws.Cells(r,15).Value=CurrentUser()
            WriteAuditLog "COMPENSATE","Card_Master",cardID,"RemainingSessions",CStr(curRemain),CStr(snapshotRemaining),reason
            Exit Sub
        End If
    Next r
End Sub

' Backward-compatible wrapper around Calculate+Commit. Any caller using the
' original DeductSessions(cardID, requested, shortage) signature keeps
' working unchanged and now goes through the same Balance Guard.
Public Function DeductSessions(ByVal cardID As String, ByVal requested As Double, ByRef shortage As Double) As Boolean
    Dim beforeUsed As Double, beforeRemain As Double, usedS As Double, afterRemain As Double
    shortage=0
    If Len(cardID)=0 Then Exit Function
    If Not CalculateDeduction(cardID, requested, beforeUsed, beforeRemain, usedS, shortage, afterRemain) Then Exit Function
    CommitDeduction cardID, beforeRemain, usedS, afterRemain
    DeductSessions=True
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

' PATCH-02 / P0-04: production CardID format is PREFIX + YYYYMMDD + ###
' (was PREFIX + YYMMDD + ###). Historical CardIDs (pre-existing rows using
' other formats, e.g. ROC-calendar-based IDs from paper migration) are never
' touched by this function - it only ever generates NEW IDs going forward.
Public Function GenerateCardID(ByVal prefix As String, ByVal purchaseDate As Date) As String
    Dim ws As Worksheet, r As Long, d As String, seq As Long, v As String
    Set ws=ThisWorkbook.Worksheets("Card_Master"): d=Format$(purchaseDate,"yyyymmdd")
    For r=2 To LastDataRow(ws,1)
        v=CStr(ws.Cells(r,1).Value)
        If Left$(v,Len(prefix)+8)=prefix & d Then seq=WorksheetFunction.Max(seq,Val(Right$(v,3)))
    Next r
    GenerateCardID=prefix & d & Format$(seq+1,"000")
End Function

' ============================================================
' PATCH-01 additions (unchanged in PATCH-02)
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

' ============================================================
' PATCH-02 / P1-02: CardType + audit metadata persistence.
' New Card_Master columns appended at 11-15; existing 1-10 untouched.
' Historical rows (pre-existing, CardType unknown) are never backfilled --
' see Contract V1.1 S7: historical Card_Master rows must not be guessed or
' retroactively classified.
' ============================================================
Private Sub EnsureCardMasterAuditColumns(ByVal ws As Worksheet)
    EnsureColumnHeader ws, 11, "CardType"
    EnsureColumnHeader ws, 12, "CreatedAt"
    EnsureColumnHeader ws, 13, "CreatedBy"
    EnsureColumnHeader ws, 14, "ModifiedAt"
    EnsureColumnHeader ws, 15, "ModifiedBy"
End Sub

' Writes one new card into Card_Master.
' Column order (verified against the live workbook header row, PATCH-02 appends 11-15):
' 1 CardID | 2 CustomerID | 3 PurchaseDate | 4 PurchaseAmount | 5 OriginalSessions
' 6 UsedSessions | 7 RemainingSessions | 8 PaymentMethod | 9 CardStatus | 10 Notes
' 11 CardType | 12 CreatedAt | 13 CreatedBy | 14 ModifiedAt | 15 ModifiedBy
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
    EnsureCardMasterAuditColumns ws
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
    ws.Cells(r, 11).Value = cardType
    ws.Cells(r, 12).Value = Now
    ws.Cells(r, 13).Value = CurrentUser()
    ws.Cells(r, 14).Value = Now
    ws.Cells(r, 15).Value = CurrentUser()

    WriteAuditLog "CREATE", "Card_Master", newID, "RemainingSessions", "", CStr(originalSessions), "New card purchase"
    WriteAuditLog "CREATE", "Card_Master", newID, "CardType", "", cardType, "New card purchase"
    CreateCard = newID
End Function
