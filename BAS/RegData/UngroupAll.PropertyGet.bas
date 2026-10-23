Property Get UngroupAll() As Boolean
    If GetSetting(AppName, "ConvertAllToCurves", "UngroupAll") = "1" Then UngroupAll = True
End Property
