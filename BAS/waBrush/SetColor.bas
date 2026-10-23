Public Sub SetColor(ColorRGB As Long, Opacity As Long)
    If Initialized Then Class_Terminate
    GdipCreateSolidFill ConvertColor(ColorRGB, Opacity), lPtr
    Initialized = True
End Sub
