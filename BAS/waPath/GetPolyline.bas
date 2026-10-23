Public Function GetPolyline(segment As segment, Optional ByVal flatness As Single = 1 / 4) As segment()
    Dim path As waPath
    Dim segments() As segment
    Set path = New waPath
    With segment
        path.AddBezierSegment .P1.x, .P1.y, .P2.x, .P2.y, .P3.x, .P3.y, .P4.x, .P4.y
        GdipFlattenPath path.Ptr, , flatness
        segments = path.segments
        segments(0).P2 = .P1
        segments(1).P1 = .P1
        segments(UBound(segments)).P2 = .P4
    End With
    GetPolyline = segments
End Function
