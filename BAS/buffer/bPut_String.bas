Public Function bPut_String(ByVal value As String) As Long
    If cp + Len(value) * 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim q As Long
        Dim i As Integer
        For q = 1 To Len(value)
            i = AscW(Mid(value, q, 1))
            CopyMemory my_bytes((q - 1) * 2 + cp), i, 2
        Next
        cp = cp + Len(value) * 2
        NeedCalculateCRC = True
    End If
End Function
