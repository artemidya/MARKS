Public Function bPut_Bytes(ByRef value() As Byte)
    Dim cnt As Long
    cnt = UBound(value) + 1
    If cp + cnt > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory my_bytes(cp), value(0), cnt
        cp = cp + cnt
        NeedCalculateCRC = True
    End If
End Function
