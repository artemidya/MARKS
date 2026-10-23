Public Property Let LaserDY(oValue As Long)
    If oValue > 200 Then oValue = 200
    If oValue < -200 Then oValue = -200
    If oValue <> lLaserDY Then NeedSave = True
    lLaserDY = oValue
End Property
