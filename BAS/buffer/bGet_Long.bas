Public Function bGet_Long() As Long
    If cp + 4 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory bGet_Long, my_bytes(cp), 4
        If InterPlatformNetConverter Then bGet_Long = ntohl(bGet_Long)
        cp = cp + 4
    End If
End Function
