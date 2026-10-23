Private Function FolderExists(ByVal FolderPath As String) As Boolean
    On Error GoTo 10
    If right(FolderPath, 1) = "\" Then
        FolderPath = left(FolderPath, Len(FolderPath) - 1)
    End If
    
    FolderExists = (Dir(FolderPath, vbDirectory) <> "")
10
End Function
