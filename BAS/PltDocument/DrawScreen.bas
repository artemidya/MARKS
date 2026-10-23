Private Sub DrawScreen()
    myGraphics.ResetClip
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.FillRect lBackgroundBrush, -1, -1, ScreenWidth + 1, ScreenHeight + 1
    If lViewBorder Then myGraphics.DrawRect CreatePen(lBorderColor), 0, 0, ScreenWidth - 1, ScreenHeight - 1
End Sub
