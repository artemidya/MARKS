Property Let ZeroPosition(oValue As ZeroPositions)
    lZeroPosition = oValue
    
    'Init Plotter
    If PlotterModelString = "" Then
        If MaxSheetWidth > 600000 Then
            ActivePlotter = GetModel_1600
        Else
            ActivePlotter = GetModel_600
        End If
    Else
        ActivePlotter = GetPlotterFromString(PlotterModelString, PlotterName)
        ActivePlotter.PlotterWorkWidth = MaxSheetWidth
    End If
    
    CalculatePage
    Sorting
End Property
