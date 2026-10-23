Property Get ReduceCurves() As Boolean
    If GetSetting(AppName, "Regroup", "ReduceCurves") = "1" Then ReduceCurves = True
End Property
