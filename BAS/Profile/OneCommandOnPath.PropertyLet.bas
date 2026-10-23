Public Property Let OneCommandOnPath(oValue As Boolean)
    If oValue <> lOneCommandOnPath Then NeedSave = True
    lOneCommandOnPath = oValue
End Property
