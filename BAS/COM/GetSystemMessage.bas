Public Function GetSystemMessage(lngErrorCode As Long) As String
Dim intPos As Integer
Dim strMessage As String, strMsgBuff As String * 256

    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngErrorCode, 0, strMsgBuff, 255, 0)

    intPos = InStr(1, strMsgBuff, vbNullChar)
    If intPos > 0 Then
        strMessage = Trim$(left(strMsgBuff, intPos - 1))
    Else
        strMessage = Trim$(strMsgBuff)
    End If
    
    GetSystemMessage = strMessage
    
End Function
