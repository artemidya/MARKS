Property Get FirstPoint() As POINTL
    Dim segments() As segment
    segments = Me.segments
    FirstPoint = segments(1).P1
End Property
