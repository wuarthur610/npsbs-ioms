Attribute VB_Name = "modText"
Option Explicit

' ASCII-only source. Builds Traditional Chinese at runtime via Unicode code points.
Public Function UText(ParamArray cp() As Variant) As String
    Dim i As Long, s As String
    For i = LBound(cp) To UBound(cp)
        s = s & ChrW$(CLng(cp(i)))
    Next i
    UText = s
End Function

Public Function MsgText(ByVal key As String) As String
    Select Case UCase$(key)
        Case "DUP_CUSTOMER": MsgText = UText(&H540C&, &H540D&, &H5BA2&, &H6236&, &H8D85&, &H904E&, &H4E00&, &H7B46&, &HFF0C&, &H8ACB&, &H4EBA&, &H5DE5&, &H78BA&, &H8A8D&, &H2048&)
        Case "CUSTOMER_REQUIRED": MsgText = UText(&H8ACB&, &H8F38&, &H5165&, &H5BA2&, &H6236&, &H59D3&, &H540D&, &H3002&)
        Case "CUSTOMER_EXISTS": MsgText = UText(&H5BA2&, &H6236&, &H5DF2&, &H5B58&, &H5728&, &HFF0C&, &H8ACB&, &H76F4&, &H63A5&, &H9078&, &H53D6&, &H65E2&, &H6709&, &H5BA2&, &H6236&, &H3002&)
        Case "CUSTOMER_NOT_FOUND": MsgText = UText(&H627E&, &H4E0D&, &H5230&, &H6307&, &H5B9A&, &H5BA2&, &H6236&, &H3002&)
        Case "SESSION_REQUIRED": MsgText = UText(&H5802&, &H6578&, &H5FC5&, &H9808&, &H5927&, &H65BC&, &H3000&, &H3002&)
        Case "THERAPIST_REQUIRED": MsgText = UText(&H8ACB&, &H9078&, &H64C7&, &H8ABF&, &H7406&, &H5E2B&, &H3002&)
        Case "DATE_REQUIRED": MsgText = UText(&H8ACB&, &H8F38&, &H5165&, &H8ABF&, &H7406&, &H65E5&, &H671F&, &H3002&)
        Case "END_BEFORE_START": MsgText = UText(&H7D50&, &H675F&, &H6642&, &H9593&, &H4E0D&, &H53EF&, &H65E9&, &H65BC&, &H958B&, &H59CB&, &H6642&, &H9593&, &H3002&)
        Case "NO_CUSTOMER": MsgText = UText(&H627E&, &H4E0D&, &H5230&, &H5BA2&, &H6236&, &HFF0C&, &H8ACB&, &H5148&, &H5EFA&, &H7ACB&, &H5BA2&, &H6236&, &H3002&)
        Case "NO_CARD": MsgText = UText(&H627E&, &H4E0D&, &H5230&, &H6307&, &H5B9A&, &H5361&, &H7247&, &H3002&)
        Case "CARD_CUSTOMER": MsgText = UText(&H6B64&, &H5361&, &H7247&, &H4E0D&, &H5C6C&, &H65BC&, &H76EE&, &H524D&, &H5BA2&, &H6236&, &H3002&)
        Case "CARD_DATE": MsgText = UText(&H5957&, &H5361&, &H8CFC&, &H8CB7&, &H65E5&, &H671F&, &H665A&, &H65BC&, &H6CBB&, &H7642&, &H65E5&, &H671F&, &HFF0C&, &H4E0D&, &H5F97&, &H56DE&, &H6263&, &H3002&)
        Case "NO_NEGATIVE": MsgText = UText(&H5957&, &H5361&, &H5269&, &H9918&, &H5802&, &H6578&, &H4E0D&, &H5F97&, &H5C0F&, &H65BC&, &H96F6&, &H3002&)
        Case "SHORTAGE_RESOLUTION": MsgText = UText(&H4E0D&, &H8DB3&, &H5802&, &H6578&, &H5FC5&, &H9808&, &H9078&, &H64C7&, &H8655&, &H7406&, &H65B9&, &H5F0F&, &H3002&)
        Case "SHORTAGE_FEE": MsgText = UText(&H4E0D&, &H8DB3&, &H5802&, &H6578&, &H6536&, &H8CBB&, &H6642&, &HFF0C&, &H8ACB&, &H8F38&, &H5165&, &H6536&, &H8CBB&, &H91D1&, &H984D&, &H3002&)
        Case "PENDING_CLOSED": MsgText = UText(&H6B64&, &H5F85&, &H88DC&, &H6263&, &H5DF2&, &H7D50&, &H6848&, &H3002&)
        Case "NEW_CARD_SHORT": MsgText = UText(&H65B0&, &H5361&, &H5802&, &H6578&, &H4E0D&, &H8DB3&, &HFF0C&, &H4E0D&, &H5141&, &H8A31&, &H8CA0&, &H5802&, &H6578&, &H3002&)
        Case "INV_REASON": MsgText = UText(&H76E4&, &H9EDE&, &H6709&, &H5DEE&, &H7570&, &H6642&, &H5FC5&, &H9808&, &H9078&, &H64C7&, &H5DEE&, &H7570&, &H539F&, &H56E0&, &H3002&)
        Case "INIT_OK": MsgText = UText(&H7CFB&, &H7D71&, &H521D&, &H59CB&, &H5316&, &H5B8C&, &H6210&, &H3002&)
        Case "SMOKE_PASS": MsgText = "Production Smoke Test: PASS"
        Case "SMOKE_FAIL": MsgText = "Production Smoke Test: FAIL"
        Case "PENDING_NOT_FOUND": MsgText = UText(&H627E&, &H4E0D&, &H5230&, &H6307&, &H5B9A&, &H7684&, &H5F85&, &H88DC&, &H6263&, &H7D00&, &H9304&, &H3002&)
        Case "SHORTAGE_ALREADY_RESOLVED": MsgText = UText(&H6B64&, &H7B46&, &H8ABF&, &H7406&, &H7684&, &H4E0D&, &H8DB3&, &H5802&, &H6578&, &H5DF2&, &H8655&, &H7406&, &H904E&, &HFF0C&, &H4E0D&, &H53EF&, &H91CD&, &H8907&, &H8655&, &H7406&, &H3002&)
        Case "NO_SHORTAGE": MsgText = UText(&H6B64&, &H7B46&, &H8ABF&, &H7406&, &H6C92&, &H6709&, &H4E0D&, &H8DB3&, &H5802&, &H6578&, &HFF0C&, &H7121&, &H9808&, &H8655&, &H7406&, &H3002&)
        Case "CARD_PREFIX_UNCONFIRMED": MsgText = UText(&H6620&, &H904A&, &H5361&, &H7684&, &H5361&, &H865F&, &H524D&, &H7DB4&, &H5C1A&, &H672A&, &H78BA&, &H8A8D&, &HFF0C&, &H8ACB&, &H5148&, &H78BA&, &H8A8D&, &H5F8C&, &H518D&, &H5EFA&, &H7ACB&, &H3002&)
        Case "CARD_TYPE_UNKNOWN": MsgText = UText(&H672A&, &H77E5&, &H7684&, &H5361&, &H7A2E&, &HFF0C&, &H7121&, &H6CD5&, &H7522&, &H751F&, &H5361&, &H865F&, &H3002&)
        Case "AMOUNT_NEGATIVE": MsgText = UText(&H8CFC&, &H8CB7&, &H91D1&, &H984D&, &H4E0D&, &H5F97&, &H70BA&, &H8CA0&, &H6578&, &H3002&)
        Case "CARD_ID_DUPLICATE": MsgText = UText(&H5361&, &H865F&, &H5DF2&, &H5B58&, &H5728&, &HFF0C&, &H8ACB&, &H78BA&, &H8A8D&, &H3002&)
        Case "CARD_TYPE_REQUIRED": MsgText = UText(&H8ACB&, &H9078&, &H64C7&, &H5361&, &H7A2E&, &H3002&)
        Case "BALANCE_GUARD": MsgText = UText(&H6263&, &H5802&, &H5F8C&, &H9918&, &H984D&, &H4E0D&, &H53EF&, &H5C0F&, &H65BC&, &H96F6&, &HFF0C&, &H4EA4&, &H6613&, &H5DF2&, &H62D2&, &H7D55&, &H3002&)
        Case "SETTLEMENT_FAILED": MsgText = UText(&H88DC&, &H6263&, &H7D50&, &H7B97&, &H5931&, &H6557&, &HFF0C&, &H5DF2&, &H9084&, &H539F&, &H5957&, &H5361&, &H9918&, &H984D&, &H4E26&, &H4FDD&, &H7559&, &H5F85&, &H88DC&, &H6263&, &H72C0&, &H614B&, &H3002&)
        Case Else: MsgText = key
    End Select
End Function

Public Function TW_Cash() As String
    TW_Cash = UText(&H73FE&, &H91D1&)
End Function
Public Function TW_Bank() As String
    TW_Bank = UText(&H9280&, &H884C&, &H8F49&, &H5E33&)
End Function
Public Function TW_Charge() As String
    TW_Charge = UText(&H6536&, &H53D6&, &H6E1B&, &H5C11&, &H5802&, &H6578&, &H8CBB&, &H7528&)
End Function
Public Function TW_Waive() As String
    TW_Waive = UText(&H76F4&, &H63A5&, &H6E1B&, &H514D&)
End Function
Public Function TW_Pending() As String
    TW_Pending = UText(&H66AB&, &H6642&, &H4FDD&, &H7559&, &H81F3&, &H4E0B&, &H6B21&, &H8CFC&, &H8CB7&, &H518D&, &H6263&, &H5802&, &H6578&)
End Function
Public Function TW_Completed() As String
    TW_Completed = UText(&H5DF2&, &H5B8C&, &H6210&)
End Function
Public Function TW_NoCard() As String
    TW_NoCard = UText(&H7121&, &H5361&)
End Function
Public Function TW_AutoDeducted() As String
    TW_AutoDeducted = UText(&H5DF2&, &H81EA&, &H52D5&, &H6263&, &H5802&)
End Function
Public Function TW_Insufficient() As String
    TW_Insufficient = UText(&H5802&, &H6578&, &H4E0D&, &H8DB3&)
End Function

' PATCH-01: card type display names (ASCII-safe, built at runtime)
Public Function TW_CardCaimian() As String
    TW_CardCaimian = UText(&H5F69&, &H7DBF&, &H5361&)
End Function

Public Function TW_CardHuaxia() As String
    TW_CardHuaxia = UText(&H83EF&, &H590F&, &H5C0A&, &H76C8&, &H5361&)
End Function

Public Function TW_CardHejia() As String
    TW_CardHejia = UText(&H95D4&, &H5BB6&, &H842C&, &H5409&, &H5361&)
End Function

Public Function TW_CardYingyou() As String
    TW_CardYingyou = UText(&H6620&, &H904A&, &H5361&)
End Function
