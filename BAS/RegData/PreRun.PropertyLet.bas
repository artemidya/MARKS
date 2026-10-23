Property Let PreRun(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, Section, "PreRun", "1"
    Else
        SaveSetting AppName, Section, "PreRun", "0"
    End If
End Property
