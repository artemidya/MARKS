Property Let ReduceCurves(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Regroup", "ReduceCurves", "1"
    Else
        SaveSetting AppName, "Regroup", "ReduceCurves", "0"
    End If
End Property
