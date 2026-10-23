Private Sub DrawRect(ByVal mPen As LongPtr, ByVal mx As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long)
    GdipDrawRectangleI HGraphics, mPen, shiftX + mx, shiftY + mY, mWidth, mHeight
End Sub
