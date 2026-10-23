Property Let Initializer(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lInitializer Then NeedSave = True
    lInitializer = Trim(oValue)
End Property
