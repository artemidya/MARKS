Public Function bGet_String(lenSymbols As Long) As String
    If cp + lenSymbols * 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim q As Long
        Dim i As Integer
        For q = 1 To lenSymbols
            CopyMemory i, my_bytes((q - 1) * 2 + cp), 2
            bGet_String = bGet_String & ChrW(i)
        Next
        cp = cp + lenSymbols * 2
    End If
End Function
