Public Function GetModel_1600() As Plotter2DModel
    If Not StandardPlotter1600_2DModel_Initialized Then
        InitStandardPlotter1600_2DModel
        
        With StandardPlotter1600_2DModel
            Dim boundingBox As RECTL
            boundingBox = CalculatePlotterBoundingBox(.Curves)
            Set .CarriageCurves = GetStandardCarriageCurves
            .PlotterWorkWidth = 1600000
            .top = boundingBox.top
            .bottom = boundingBox.bottom
            .left = boundingBox.left
            .right = boundingBox.right
            .Height = .top - .bottom
            .Width = .right - .left
        End With
        
    End If
    GetModel_1600 = StandardPlotter1600_2DModel
End Function
