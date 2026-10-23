Property Get FirstCutDateLng() As Long
    On Error GoTo 10
    FirstCutDateLng = CLng(GetSetting(AppName, Section, "FirstCutDateLng"))
10
End Property
