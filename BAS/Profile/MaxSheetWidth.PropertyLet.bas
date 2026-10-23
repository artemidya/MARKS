Property Let MaxSheetWidth(oValue As Long)
    If oValue <> lMaxSheetWidth Then NeedSave = True
    lMaxSheetWidth = oValue
End Property
