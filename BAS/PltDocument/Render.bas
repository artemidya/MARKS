Public Sub Render()
    Static t As Single
    t = Timer
    DrawScreen
    DrawPlotterShadow
    DrawSheetShadow
    DrawPlotter
    DrawCarriage
    DrawSheet
    DrawCurves
    DrawLogo
    DrawRedLines
    DrawTopMaterialShadow
    DrawAxis
    Refresh
    RaiseEvent RenderTime(Timer - t)
End Sub
