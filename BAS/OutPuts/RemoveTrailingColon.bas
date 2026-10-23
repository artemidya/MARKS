Function RemoveTrailingColon(ByVal s As String) As String
    If right(s, 1) = ":" Then
        RemoveTrailingColon = left(s, Len(s) - 1)
    Else
        RemoveTrailingColon = s
    End If
End Function
