Private Function GetMoneyUnit() As String
    On Error GoTo 10
    GetMoneyUnit = "$"
    Dim buffer As String, ret As Long
    buffer = String(256, 0)
    ret = GetLocaleInfo(1024, &H14, buffer, Len(buffer))
    If ret > 0 Then
        GetMoneyUnit = left(buffer, ret - 1)
    End If
    Exit Function
10:
End Function
