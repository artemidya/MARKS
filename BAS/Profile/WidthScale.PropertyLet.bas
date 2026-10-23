Public Property Let WidthScale(oValue As Double)
    If oValue = 0 Then oValue = 1
    If oValue > 5 Then oValue = 5
    If oValue < -5 Then oValue = -5
    If oValue <> lWidthScale Then NeedSave = True
    lWidthScale = oValue
End Property
