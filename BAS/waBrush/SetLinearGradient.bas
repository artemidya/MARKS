Public Sub SetLinearGradient(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ColorRGB1 As Long, ColorRGB2 As Long, Optional WrapMode As WrapMode = WrapModeTileFlipXY, Optional Opacity1 As Long = 100, Optional Opacity2 As Long = 100)
    If Initialized Then Class_Terminate
    
    P1.x = x1
    P1.y = y1
    P2.x = x2
    P2.y = y2
    
    GdipCreateLineBrushI P1, P2, ConvertColor(ColorRGB1, Opacity1), ConvertColor(ColorRGB2, Opacity2), WrapMode, lPtr
    Initialized = True
End Sub
