Private Sub DrawLine(ByVal mPen As LongPtr, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    GdipDrawLineI HGraphics, mPen, shiftX + x1, shiftY + y1, shiftX + x2, shiftY + y2
End Sub
