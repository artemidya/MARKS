Property Get TrimOffset() As Long
    On Error GoTo 10
    TrimOffset = CLng(GetSetting(AppName, "Trim", "TrimOffset"))
    If TrimOffset < 0 Then TrimOffset = 0
10:
End Property
