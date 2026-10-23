Public Function StrToAscii(str As String) As Byte()
    Dim l As Long: l = Len(str)
    If l > 0 Then
        Dim bytes() As Byte
        ReDim bytes(l - 1)
        Dim q As Long
        For q = 0 To l - 1
            bytes(q) = Asc(Mid(str, q + 1, 1))
        Next
        StrToAscii = bytes
    End If
End Function
