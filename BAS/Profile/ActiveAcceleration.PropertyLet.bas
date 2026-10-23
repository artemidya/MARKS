Property Let ActiveAcceleration(oValue As Long)
    If oValue <> lActiveAcceleration Then NeedSave = True
    lActiveAcceleration = oValue
End Property
