Property Let ActiveSpeed(oValue As Long)
    If oValue <> lActiveSpeed Then NeedSave = True
    lActiveSpeed = oValue
End Property
