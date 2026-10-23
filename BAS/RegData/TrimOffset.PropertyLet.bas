Property Let TrimOffset(oValue As Long)
    If oValue < 0 Then oValue = 0
    SaveSetting AppName, "Trim", "TrimOffset", oValue
End Property
