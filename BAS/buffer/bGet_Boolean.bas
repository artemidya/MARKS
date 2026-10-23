Public Function bGet_Boolean() As Boolean
    If cp + LenB(bGet_Boolean) > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory bGet_Boolean, my_bytes(cp), LenB(bGet_Boolean)
        cp = cp + LenB(bGet_Boolean)
    End If
End Function
