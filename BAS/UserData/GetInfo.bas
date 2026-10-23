Private Function GetInfo() As String
    On Error GoTo 10
    Dim buffer As String, ret As Long
    buffer = String(256, 0)
    ret = GetLocaleInfo(1024, &H5C, buffer, Len(buffer))
    If ret > 1 Then
        GetInfo = left(buffer, 2)
    Else
        GetInfo = ""
    End If
    Exit Function
10:
End Function
