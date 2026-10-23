Private Sub TerminateGDI()
    GdipDeletePen mPen
    GdipDeletePen LinePen
    GdipDeletePen MaterialPen
    GdipDeletePen PagesPen
    GdipDeletePen FieldsPen
    
    GdipDeletePen LinesPaperPen
    GdipDeletePen LinesShadowPen
    
    GdipDeleteBrush hBrush
    GdipDeleteBrush PaperBrush
    GdipDeleteBrush ShadowBrush
    
    
    GdipDeleteGraphics HGraphics
    GdiplusShutdown GdipToken
    ReleaseDC hWndForm, HDCForm
End Sub
