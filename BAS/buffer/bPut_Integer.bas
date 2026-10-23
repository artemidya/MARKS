Public Function bPut_Integer(ByVal value As Integer) As Long
    If cp + 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        If InterPlatformNetConverter Then value = htons(value)
        CopyMemory my_bytes(cp), value, 2
        cp = cp + 2
        NeedCalculateCRC = True
    End If
End Function
