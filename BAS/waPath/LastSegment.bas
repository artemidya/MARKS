Property Get LastSegment() As segment
    Dim segments() As segment
    segments = Me.segments
    LastSegment = segments(UBound(segments))
End Property
