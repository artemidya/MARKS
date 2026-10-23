Public Property Let HeightScale(oValue As Double)
    If oValue = 0 Then oValue = 1
    If oValue > 5 Then oValue = 5
    If oValue < -5 Then oValue = -5
    If oValue <> lHeightScale Then NeedSave = True
    lHeightScale = oValue
End Property
