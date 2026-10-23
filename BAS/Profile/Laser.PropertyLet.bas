Public Property Let Laser(oValue As Boolean)
    If oValue <> lLaser Then NeedSave = True
    lLaser = oValue
End Property
