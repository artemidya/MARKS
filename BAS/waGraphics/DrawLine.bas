Public Sub DrawLine(pen As waPen, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    GdipDrawLineI lHGraphics, pen.Ptr, x1, y1, x2, y2
End Sub
