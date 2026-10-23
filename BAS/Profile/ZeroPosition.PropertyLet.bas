Public Property Let ZeroPosition(oValue As ZeroPositions)
    If oValue > 2 Or oValue < 0 Then oValue = 0
    If oValue <> lZeroPosition Then NeedSave = True
    lZeroPosition = oValue
End Property
