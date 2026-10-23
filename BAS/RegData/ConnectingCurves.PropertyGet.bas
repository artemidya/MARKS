Property Get ConnectingCurves() As Boolean
    If GetSetting(AppName, "Regroup", "ConnectingCurves") = "1" Then ConnectingCurves = True
End Property
