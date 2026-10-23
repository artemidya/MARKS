Public Sub SetScreenSize(ScreenWidth As Long, ScreenHeight As Long)
    Me.ScreenWidth = ScreenWidth
    Me.ScreenHeight = ScreenHeight
    MemoryBitmap1.SetSize ScreenWidth, ScreenHeight
    myGraphics.hDC = MemoryBitmap1.hDC
End Sub
