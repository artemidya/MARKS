Property Let UngroupAll(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "ConvertAllToCurves", "UngroupAll", "1"
    Else
        SaveSetting AppName, "ConvertAllToCurves", "UngroupAll", "0"
    End If
End Property
