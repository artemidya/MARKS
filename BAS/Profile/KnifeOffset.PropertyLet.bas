Public Property Let KnifeOffset(oValue As Long)
    If oValue > 5550 Then oValue = 5550
    If oValue < 0 Then oValue = 0
    oValue = (oValue \ 25) * 25
    If oValue <> lKnifeOffset Then NeedSave = True
    lKnifeOffset = oValue
End Property
