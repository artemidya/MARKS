Property Get PreRun() As Boolean
    If GetSetting(AppName, Section, "PreRun") = "1" Then PreRun = True
End Property
