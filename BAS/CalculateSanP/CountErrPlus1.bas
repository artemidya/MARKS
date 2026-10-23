Private Sub CountErrPlus1(ByRef shape As VGCore.shape)
    CountErr = CountErr + 1
    ReDim Preserve ErrOnScreenCurves(CountErr)
    ReDim Preserve ErrOnScreenText(CountErr)
    
    Set ErrOnScreenText(CountErr) = New OnScreenText
    ErrOnScreenText(CountErr).SetTextColor RGB(250, 0, 0)
    ErrOnScreenText(CountErr).SetPixelOffset 15, -10
    ErrOnScreenText(CountErr).SetTextAndPosition "ERR-" & CountErr, shape.PositionX, shape.PositionY, 2 'cdrCenterOnScreenText ' cdrCenterOnSreenText
    'ErrOnScreenText(CountErr).Show

    Set ErrOnScreenCurves(CountErr) = New OnScreenCurve
    With shape.boundingBox
        ErrOnScreenCurves(CountErr).SetPen RGB(250, 0, 0), 1, cdrOnScreenCurvePenSolid
        Dim curve As curve
        Set curve = CreateErrCurve(.x, .y, .right, .top)
        ErrOnScreenCurves(CountErr).SetCurve curve
        'ErrOnScreenCurves(CountErr).Show
    End With
End Sub
