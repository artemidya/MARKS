Property Let MaxSheetHeight(oValue As Long)
    If oValue <> lMaxSheetHeight Then NeedSave = True
    lMaxSheetHeight = oValue
End Property
