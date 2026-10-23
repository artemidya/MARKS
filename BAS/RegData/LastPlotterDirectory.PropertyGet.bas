Property Get LastPlotterDirectory() As String
    On Error GoTo 10
    LastPlotterDirectory = GetSetting(AppName, Section, "LastPlotterDirectory")
10
End Property
