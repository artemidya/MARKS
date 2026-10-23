Public Sub Refresh()
    BitBlt hDC, PositionX, PositionY, MemoryBitmap1.Width, MemoryBitmap1.Height, MemoryBitmap1.hDC, 0, 0, SRCCOPY
End Sub
