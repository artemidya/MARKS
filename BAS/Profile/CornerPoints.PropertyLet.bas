Public Property Let CornerPoints(oValue As Boolean)
    If oValue <> lCornerPoints Then NeedSave = True
    lCornerPoints = oValue
End Property
