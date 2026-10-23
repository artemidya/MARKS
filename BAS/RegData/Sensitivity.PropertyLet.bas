Property Let Sensitivity(oValue As Long)
    If oValue < 25 Then oValue = 25
    If oValue > 10000 Then oValue = 10000
    SaveSetting AppName, "Duplicates", "Sensitivity", oValue
End Property
