Private Function CalculateSize(ByRef RPaths() As RPath) As Long
    CalculateSize = 12
    Dim q1 As Long
    For q1 = 1 To UBound(RPaths)
        CalculateSize = CalculateSize + 4 + UBound(RPaths(q1).RSegments) * 24
    Next
End Function
