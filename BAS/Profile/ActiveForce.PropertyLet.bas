Property Let ActiveForce(oValue As Long)
    If oValue <> lActiveForce Then NeedSave = True
    lActiveForce = oValue
End Property
