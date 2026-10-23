Property Let AutoUpdate(oValue As Boolean)
    If oValue Then
        SaveSetting "PlotCalc", "Settings", "AutoUpdate", "1"
    Else
        SaveSetting "PlotCalc", "Settings", "AutoUpdate", "0"
    End If
End Property
