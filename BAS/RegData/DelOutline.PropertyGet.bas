Property Get DelOutline() As Boolean
    If GetSetting(AppName, "ConvertAllToCurves", "DelOutline") = "1" Then DelOutline = True
End Property
