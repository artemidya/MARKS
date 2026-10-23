Public Function bGet_Byte() As Byte
    If cp + 1 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        bGet_Byte = my_bytes(cp)
        cp = cp + 1
    End If
End Function
