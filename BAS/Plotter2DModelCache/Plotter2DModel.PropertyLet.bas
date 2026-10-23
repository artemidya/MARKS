Property Let Plotter2DModel(oValue As Plotter2DModel)
    With oValue
        Set Curves = .Curves
        Set CarriageCurves = .CarriageCurves
        PlotterWorkWidth = .PlotterWorkWidth
        top = .top
        bottom = .bottom
        left = .left
        right = .right
        Height = .Height
        Width = .Width
    End With
End Property
