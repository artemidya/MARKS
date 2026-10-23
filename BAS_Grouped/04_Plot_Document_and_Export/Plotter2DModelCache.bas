' =============================================================================
' Group: 04_Plot_Document_and_Export
' Source module: Plotter2DModelCache.cls
' Procedures: 2
' =============================================================================

' -----------------------------------------------------------------------------
' Plotter2DModel (PropertyGet), source lines 21-33
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' Plotter2DModel (PropertyLet), source lines 35-47
' -----------------------------------------------------------------------------
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
