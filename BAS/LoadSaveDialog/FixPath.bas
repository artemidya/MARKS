Private Function FixPath(ByVal path As String) As String
    If right(path, 1) <> "\" Then
        FixPath = path & "\"
    Else
        FixPath = path
    End If
End Function
