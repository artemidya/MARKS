Property Let DelFill(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "ConvertAllToCurves", "DelFill", "1"
    Else
        SaveSetting AppName, "ConvertAllToCurves", "DelFill", "0"
    End If
End Property
