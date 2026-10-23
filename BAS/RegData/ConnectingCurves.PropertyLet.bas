Property Let ConnectingCurves(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Regroup", "ConnectingCurves", "1"
    Else
        SaveSetting AppName, "Regroup", "ConnectingCurves", "0"
    End If
End Property
