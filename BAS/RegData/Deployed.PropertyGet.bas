Property Get Deployed() As Boolean
    If GetSetting(AppName, Section, "Deployed") = "1" Then Deployed = True
End Property
