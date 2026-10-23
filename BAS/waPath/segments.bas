Property Get segments() As segment()
    
    Dim count As Long
    GdipGetPointCount Ptr, count
    Dim p() As POINTL
    Dim t() As Byte
    
    ReDim p(count - 1)
    ReDim t(count - 1)
    GdipGetPathPointsI Ptr, p(0), count
    GdipGetPathTypes Ptr, t(0), count
    

    
    Dim cps As Long: cps = -1
    Dim q As Long
    Dim LastPoint As POINTL
    
    Dim countsegments As Long
    For q = 0 To count - 1
        Select Case t(q)
            Case 0  'Start
                countsegments = countsegments + 1
            Case 1  'Line
                countsegments = countsegments + 1
            Case 129 'Line (end)
                countsegments = countsegments + 1
            Case 3  'Bezier
                countsegments = countsegments + 1
                q = q + 2
        End Select
    Next
    
    Dim s() As segment
    ReDim s(countsegments - 1)
    
    For q = 0 To count - 1
        Select Case t(q)
            Case 0  'Start
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).Type = InvisibleLine
                LastPoint = s(cps).P2
            Case 1  'Line
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).Type = LineSegment
                LastPoint = s(cps).P2
            Case 129 'Line (end)
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).Type = LineSegment
                LastPoint = s(cps).P2
            Case 3  'Bezier
                cps = cps + 1
                s(cps).P1 = LastPoint
                s(cps).P2 = p(q)
                s(cps).P3 = p(q + 1)
                s(cps).P4 = p(q + 2)
                LastPoint = s(cps).P4
                q = q + 2
                s(cps).Type = CurveSegment
                
        End Select
    Next

    segments = s
End Property
