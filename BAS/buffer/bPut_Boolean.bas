Public Function bPut_Boolean(ByVal value As Boolean) As Long
    If cp + LenB(value) > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory my_bytes(cp), value, LenB(value)
        cp = cp + LenB(value)
        NeedCalculateCRC = True
    End If
End Function
