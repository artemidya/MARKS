Property Let Deployed(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, Section, "Deployed", "1"
    Else
        SaveSetting AppName, Section, "Deployed", "0"
    End If
End Property
