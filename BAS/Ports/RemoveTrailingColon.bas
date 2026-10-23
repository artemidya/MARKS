Private Function RemoveTrailingColon(ByVal s As String) As String
    On Error GoTo 10
    If right(s, 1) = ":" Then
        RemoveTrailingColon = left(s, Len(s) - 1)
    Exit Function
    End If
10
    RemoveTrailingColon = s
End Function
