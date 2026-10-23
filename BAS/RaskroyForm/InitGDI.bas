Private Sub InitGDI()
    ' Get hWndForm
    hWndForm = FindWindow(vbNullString, Me.Caption)
    HDCForm = GetDC(hWndForm)
    
    Dim gdipStartupInput As GDIPlusStartupInput
    gdipStartupInput.GdiPlusVersion = 1&
    GdiplusStartup GdipToken, gdipStartupInput, ByVal 0
    GdipCreateFromHDC HDCForm, HGraphics
    GdipCreatePen1 ConvertColor(vbBlack, 100), 1, 2, mPen
    GdipCreatePen1 ConvertColor(vbBlack, 50), 1, 2, LinePen
    GdipCreatePen1 ConvertColor(RGB(0, 0, 150), 100), 1, 2, MaterialPen
    GdipCreatePen1 ConvertColor(RGB(255, 98, 0), 100), 1, 2, PagesPen
    GdipCreatePen1 ConvertColor(RGB(255, 173, 122), 60), 1, 2, FieldsPen
    GdipCreatePen1 ConvertColor(vbWhite, 100), 1, 2, LinesPaperPen
    GdipCreatePen1 ConvertColor(vbBlack, 30), 1, 2, LinesShadowPen
    GdipCreateSolidFill ConvertColor(vbWhite, 100), PaperBrush
    GdipCreateSolidFill ConvertColor(RGB(205, 205, 205), 100), hBrush
    GdipCreateSolidFill ConvertColor(vbBlack, 30), ShadowBrush
    GdipSetSmoothingMode HGraphics, SmoothingModeAntiAlias
End Sub
