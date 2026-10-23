Property Let DelOutline(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "ConvertAllToCurves", "DelOutline", "1"
    Else
        SaveSetting AppName, "ConvertAllToCurves", "DelOutline", "0"
    End If
End Property
