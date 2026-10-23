Property Let BeginWorkFile(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lBeginWorkFile Then NeedSave = True
    lBeginWorkFile = Trim(oValue)
End Property
