Private Sub Class_Initialize()
    'FillModeWinding
    GdipCreatePath FillModeAlternate, Ptr

    Set lPen = DefPen
    Set lBrush = DefBrush
    NeedRecalculateBounds = True
End Sub
