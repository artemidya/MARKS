Public Property Let CoordinateSystem(oValue As CoordinateSystems)
    If oValue <> lCoordinateSystem Then NeedSave = True
    lCoordinateSystem = oValue
End Property
