Private Sub Cls(Optional NeedRepaint As Boolean)
    If NeedRepaint Then Me.Repaint
    GdipFillRectangleI HGraphics, hBrush, shiftX, shiftY, WindowWidth, WindowHeigth
    GdipDrawRectangleI HGraphics, mPen, shiftX, shiftY, WindowWidth, WindowHeigth
End Sub
