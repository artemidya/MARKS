Property Let Visible(oValue As Boolean)
    If oValue <> lVisible Then NeedSave = True
    lVisible = oValue
End Property
