Private Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemMove ByVal retString, ByVal lpStr, nStr
    retString = left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function
