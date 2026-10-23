Public Function bGet_Bytes(cnt As Long) As Byte()
    If cp + cnt > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim b() As Byte
        ReDim b(cnt - 1)
        CopyMemory b(0), my_bytes(cp), cnt
        bGet_Bytes = b
        cp = cp + cnt
    End If
End Function
