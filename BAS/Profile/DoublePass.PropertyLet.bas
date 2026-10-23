Public Property Let DoublePass(oValue As Boolean)
    If oValue <> lDoublePass Then NeedSave = True
    lDoublePass = oValue
End Property
