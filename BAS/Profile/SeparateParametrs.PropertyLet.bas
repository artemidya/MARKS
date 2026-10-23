Public Property Let SeparateParametrs(oValue As Boolean)
    If oValue <> lSeparateParametrs Then NeedSave = True
    lSeparateParametrs = oValue
End Property
