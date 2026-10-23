Private Function fFileLen(fileName As String) As Long
    On Error GoTo 10
    fFileLen = FileLen(fileName)
    Exit Function
10
End Function
