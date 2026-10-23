Property Get DelFill() As Boolean
    If GetSetting(AppName, "ConvertAllToCurves", "DelFill") = "1" Then DelFill = True
End Property
