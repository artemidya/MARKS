Private Function TrimNull(ByVal strItem As String) As String
    Dim intPos As Integer

    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        TrimNull = left(strItem, intPos - 1)
    Else
        TrimNull = strItem
    End If
End Function
