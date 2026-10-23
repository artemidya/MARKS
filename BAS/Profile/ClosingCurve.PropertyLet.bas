Public Property Let ClosingCurve(oValue As Long)
    oValue = (oValue \ 25) * 25
    If oValue > 10000 Then oValue = 10000
    If oValue < 0 Then oValue = 0
    If oValue <> lClosingCurve Then NeedSave = True
    lClosingCurve = oValue
End Property
