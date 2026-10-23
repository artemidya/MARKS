Public Function bGet_Integer() As Integer
    If cp + 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory bGet_Integer, my_bytes(cp), 2
        If InterPlatformNetConverter Then bGet_Integer = ntohs(bGet_Integer)
        cp = cp + 2
    End If
End Function
