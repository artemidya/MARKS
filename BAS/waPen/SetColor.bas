Public Sub SetColor(ColorRGB As Long, Opacity As Long, Optional Width As Single = 1)
    If Initialized Then Class_Terminate
    GdipCreatePen1 ConvertColor(ColorRGB, Opacity), Width, 2, lPtr
    Initialized = True
End Sub
