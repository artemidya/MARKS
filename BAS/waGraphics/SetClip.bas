Public Sub SetClip(x As Long, y As Long, Width As Long, Height As Long, Optional CombineMode As CombineMode = CombineModeReplace)
    GdipSetClipRectI lHGraphics, x, y, Width, Height, CombineMode
End Sub
