Property Get Plotter2DModel() As Plotter2DModel
    With Plotter2DModel
        Set .Curves = Curves
        Set .CarriageCurves = CarriageCurves
        .PlotterWorkWidth = PlotterWorkWidth
        .top = top
        .bottom = bottom
        .left = left
        .right = right
        .Height = Height
        .Width = Width
    End With
End Property
