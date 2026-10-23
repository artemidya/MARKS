Private Function StrToBite(ByRef b() As Byte) As Byte
On Error GoTo 10
Dim q As Long
    For q = 0 To UBound(b)
        b(q) = Asc(Mid(InString, pos + 1, 1))
        pos = pos + 1
    Next
10
End Function
