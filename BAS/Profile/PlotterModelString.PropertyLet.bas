Public Property Let PlotterModelString(oValue As String)
    If oValue <> lPlotterModelString Then NeedSave = True
    lPlotterModelString = oValue
End Property
