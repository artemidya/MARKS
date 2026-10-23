Private Function ShiftSegments(segments() As segment, ByVal Shift As Long) As segment()
    Dim q As Long
    For q = 1 To UBound(segments)
        segments(q) = ShiftSegment(segments(q), Shift)
    Next
    ShiftSegments = segments
End Function
