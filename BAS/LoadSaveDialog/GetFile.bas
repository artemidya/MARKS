Public Function GetFile(FilePath As String) As String
    Dim pos As Long
    pos = InStrRev(FilePath, "\")
    If pos > 0 Then
        GetFile = right(FilePath, Len(FilePath) - pos)
    End If
End Function
