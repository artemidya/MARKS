Public Property Let LineType(oValue As LineTypes)
    If oValue > 1 Or oValue < 0 Then oValue = lines
    If oValue <> lLineType Then NeedSave = True
    lLineType = oValue
End Property
