Public Property Let Device(oValue As String)
    If oValue <> lDevice Then NeedSave = True
    lDevice = oValue
End Property
