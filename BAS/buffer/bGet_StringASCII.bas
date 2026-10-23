Public Function bGet_StringASCII(lenSymbols As Long) As String
    If cp + lenSymbols > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim q As Long
        For q = 1 To lenSymbols
            bGet_StringASCII = bGet_StringASCII & Chr(cp + q - 1)
        Next
        cp = cp + lenSymbols
    End If
End Function
