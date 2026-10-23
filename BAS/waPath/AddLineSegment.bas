Public Sub AddLineSegment(x As Long, y As Long)
    Static p As POINTL
    p.x = x
    p.y = y
    GdipAddPathLine2I Ptr, p, 1
    NeedRecalculateBounds = True
End Sub
