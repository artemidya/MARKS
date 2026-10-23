Property Get ActiveProfile() As Long
    On Error GoTo 10
    ActiveProfile = val(GetSetting(AppName, Section, "ActiveProfile"))
10
End Property
