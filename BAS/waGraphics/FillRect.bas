Public Sub FillRect(brush As waBrush, x As Long, y As Long, Width As Long, Height As Long)
    GdipFillRectangleI lHGraphics, brush.Ptr, x, y, Width, Height
End Sub
