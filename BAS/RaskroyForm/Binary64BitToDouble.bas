Private Function Binary64BitToDouble(ByRef BinaryString As String) As Double
    Dim bytArray(7) As Byte
    bytArray(7) = Asc(Mid(BinaryString, 8, 1))
    bytArray(6) = Asc(Mid(BinaryString, 7, 1))
    bytArray(5) = Asc(Mid(BinaryString, 6, 1))
    bytArray(4) = Asc(Mid(BinaryString, 5, 1))
    bytArray(3) = Asc(Mid(BinaryString, 4, 1))
    bytArray(2) = Asc(Mid(BinaryString, 3, 1))
    bytArray(1) = Asc(Mid(BinaryString, 2, 1))
    bytArray(0) = Asc(Mid(BinaryString, 1, 1))
    CopyMemory Binary64BitToDouble, bytArray(0), 8
End Function
