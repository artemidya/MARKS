Public Function AsciiToString(bytes() As Byte) As String
    If Not Not bytes Then
        Dim q As Long
        For q = 0 To UBound(bytes)
            AsciiToString = AsciiToString & Chr(bytes(q))
        Next
    End If
End Function
