Public Sub DrawBezier(pen As waPen, x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long, x4 As Long, y4 As Long)
    GdipDrawBezierI lHGraphics, pen.Ptr, x1, y1, x2, y2, x3, y3, x4, y4
End Sub
