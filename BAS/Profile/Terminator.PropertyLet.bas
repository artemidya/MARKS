Property Let Terminator(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lTerminator Then NeedSave = True
    lTerminator = Trim(oValue)
End Property
