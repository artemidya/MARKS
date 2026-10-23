Private Function IsPlotCalcToolBarExists() As Boolean
    On Error GoTo 10
    IsPlotCalcToolBarExists = Application.CommandBars.Item("PlotCalc").Enabled
10
End Function
