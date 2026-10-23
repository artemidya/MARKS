Property Get NestingPlaceOnPage() As Boolean
    If GetSetting(AppName, "Nesting", "NestingPlaceOnPage") = "1" Then NestingPlaceOnPage = True
End Property
