Private Sub DrawCarriage()
    'PlotterCarriageCurves
    Dim plotterMatrix As New waMatrix

    Set myGraphics.Curves = ActivePlotter.CarriageCurves
    
    Select Case lZeroPosition
        Case RightP
            plotterMatrix.Translate -WhiteOffset - lYmargin, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth, OnScreenSheetPositionY
        Case LeftP
            plotterMatrix.Translate WhiteOffset + lXmargin, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX, OnScreenSheetPositionY
        Case CenterP
            plotterMatrix.Translate lXmargin, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth \ 2, OnScreenSheetPositionY
    End Select
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
    myGraphics.SetWorldTransform plotterMatrix
    myGraphics.Render False, True
End Sub
