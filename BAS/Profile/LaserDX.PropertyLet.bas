Public Property Let LaserDX(oValue As Long)
    If oValue > 200 Then oValue = 200
    If oValue < -200 Then oValue = -200
    If oValue <> lLaserDX Then NeedSave = True
    lLaserDX = oValue
End Property
