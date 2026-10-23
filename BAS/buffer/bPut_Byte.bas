Public Function bPut_Byte(ByVal value As Byte)
    If cp + 1 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        my_bytes(cp) = value
        cp = cp + 1
        NeedCalculateCRC = True
    End If
End Function
