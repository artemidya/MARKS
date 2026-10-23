Public Sub AddBezierSegment(x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long, x4 As Long, y4 As Long)
    Static p(3) As POINTL
    p(0).x = x1
    p(0).y = y1
    p(1).x = x2
    p(1).y = y2
    p(2).x = x3
    p(2).y = y3
    p(3).x = x4
    p(3).y = y4
    GdipAddPathBeziersI Ptr, p(0), 4
    NeedRecalculateBounds = True
End Sub
