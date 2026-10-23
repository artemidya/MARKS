Property Let OutlineUse(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Duplicates", "OutlineUse", "1"
    Else
        SaveSetting AppName, "Duplicates", "OutlineUse", "0"
    End If
End Property
