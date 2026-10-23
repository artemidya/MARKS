Public Function GetModel_600() As Plotter2DModel
    If Not StandardPlotter600_2DModel_Initialized Then
        InitStandardPlotter600_2DModel
        
        With StandardPlotter600_2DModel
            Dim boundingBox As RECTL
            boundingBox = CalculatePlotterBoundingBox(.Curves)
            Set .CarriageCurves = GetStandardCarriageCurves
            .PlotterWorkWidth = 600000
            .top = boundingBox.top
            .bottom = boundingBox.bottom
            .left = boundingBox.left
            .right = boundingBox.right
            .Height = .top - .bottom
            .Width = .right - .left
        End With
        
    End If
    GetModel_600 = StandardPlotter600_2DModel
End Function
