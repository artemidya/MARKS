Public Function GetPath(FilePath As String) As String
    Dim pos As Long
    pos = InStrRev(FilePath, "\")
    If pos > 0 Then
        GetPath = left(FilePath, pos - 1)
    End If
End Function
