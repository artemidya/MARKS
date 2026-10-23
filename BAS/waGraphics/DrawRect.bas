Public Sub DrawRect(pen As waPen, x As Long, y As Long, Width As Long, Height As Long)
    GdipDrawRectangleI lHGraphics, pen.Ptr, x, y, Width, Height
End Sub
