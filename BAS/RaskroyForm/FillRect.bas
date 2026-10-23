Private Sub FillRect(ByVal hBrush As LongPtr, ByVal mx As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long)
    GdipFillRectangleI HGraphics, hBrush, shiftX + mx, shiftY + mY, mWidth, mHeight
End Sub
