Private Sub Init()
'    lhDC = GetDC(lHWndForm)
    Dim gdipStartupInput As GDIPlusStartupInput
    gdipStartupInput.GdiPlusVersion = 1&
    GdiplusStartup lGdipToken, gdipStartupInput, ByVal 0
    GdipCreateFromHDC lhDC, lHGraphics
    GdipSetSmoothingMode lHGraphics, SmoothingModeAntiAlias
'    Matrix.ScaleM 1, -1

    Initialized = True
End Sub
