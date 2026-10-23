Private Sub DrawPlotterShadow()
    Dim plotterMatrix As New waMatrix
    Set myGraphics.Curves = ActivePlotter.Curves

    Select Case lZeroPosition
        Case RightP
            plotterMatrix.Translate -WhiteOffset, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth + lShadowDistance, OnScreenSheetPositionY + lShadowDistance
        Case LeftP
            plotterMatrix.Translate WhiteOffset + ActivePlotter.PlotterWorkWidth, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + lShadowDistance, OnScreenSheetPositionY + lShadowDistance
        Case CenterP
            plotterMatrix.Translate ActivePlotter.PlotterWorkWidth \ 2, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth \ 2 + lShadowDistance, OnScreenSheetPositionY + lShadowDistance
    End Select
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
    myGraphics.SetWorldTransform plotterMatrix
    myGraphics.Render , , CreateBrush(RGB(170, 170, 170), 100)
End Sub
