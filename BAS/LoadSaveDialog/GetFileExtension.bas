Public Function GetFileExtension(file As String) As String
    Dim pos As Long
    pos = InStrRev(file, ".")
    If pos > 0 Then
        GetFileExtension = right(file, Len(file) - pos)
    End If
End Function
