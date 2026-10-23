Property Let NestingPlaceOnPage(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Nesting", "NestingPlaceOnPage", "1"
    Else
        SaveSetting AppName, "Nesting", "NestingPlaceOnPage", "0"
    End If
End Property
