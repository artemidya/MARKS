Property Get LastPoint() As POINTL
    Dim segments() As segment
    segments = Me.segments
    With segments(UBound(segments))
        If .Type = LineSegment Then
            LastPoint = .P2
        Else
            LastPoint = .P4
        End If
    End With
End Property
