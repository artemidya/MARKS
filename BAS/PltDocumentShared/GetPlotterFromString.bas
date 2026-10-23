Public Function GetPlotterFromString(ByRef PlotterString As String, ByVal PlotterName As String) As Plotter2DModel

    Dim Plotter2DModelCacheCls As Plotter2DModelCache
    PlotterName = PlotterName & CRC32(PlotterString)

    Set Plotter2DModelCacheCls = GetItemFromPlottersCacheCol(PlotterName)
    
    If Plotter2DModelCacheCls Is Nothing Then
        Dim Convert As New ConverterStringLToPlotCalcCrv
        With GetPlotterFromString
            Set .Curves = Convert.GetPlotCalcCrv(GetTag("PlotterImg", PlotterString))
            Set .CarriageCurves = Convert.GetPlotCalcCrv(GetTag("CarriageImg", PlotterString))
            Dim boundingBox As RECTL
            boundingBox = CalculatePlotterBoundingBox(.Curves)
            .top = boundingBox.top
            .bottom = boundingBox.bottom
            .left = boundingBox.left
            .right = boundingBox.right
            .Height = .top - .bottom
            .Width = .right - .left
        End With
        Set Plotter2DModelCacheCls = New Plotter2DModelCache
        Plotter2DModelCacheCls.Plotter2DModel = GetPlotterFromString
        PlottersCacheCol.Add Plotter2DModelCacheCls, PlotterName

    Else
        GetPlotterFromString = Plotter2DModelCacheCls.Plotter2DModel
    End If

End Function
