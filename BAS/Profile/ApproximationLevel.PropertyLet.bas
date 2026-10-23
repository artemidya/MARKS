Public Property Let ApproximationLevel(oValue As ApproximationLevels)
    If oValue > 2 Or oValue < 0 Then oValue = Normal
    If oValue <> lApproximationLevel Then NeedSave = True
    lApproximationLevel = oValue
End Property
