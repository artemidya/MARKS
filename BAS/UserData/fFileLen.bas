Private Function fFileLen(file As String) As Long
    On Error GoTo 10
    fFileLen = FileLen(file)
    Exit Function
10
End Function
