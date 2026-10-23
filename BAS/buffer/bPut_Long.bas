Public Function bPut_Long(ByVal value As Long) As Long
    If cp + 4 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        If InterPlatformNetConverter Then value = htonl(value)
        CopyMemory my_bytes(cp), value, 4
        cp = cp + 4
        NeedCalculateCRC = True
    End If
End Function
