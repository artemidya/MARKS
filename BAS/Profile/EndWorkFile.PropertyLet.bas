Property Let EndWorkFile(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lEndWorkFile Then NeedSave = True
    lEndWorkFile = Trim(oValue)
End Property
