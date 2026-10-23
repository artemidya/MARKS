Property Let ActiveSpeedUP(oValue As Long)
    If oValue <> lActiveSpeedUP Then NeedSave = True
    lActiveSpeedUP = oValue
End Property
