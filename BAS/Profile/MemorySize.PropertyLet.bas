Public Property Let MemorySize(oValue As Long)
    If oValue < 10000 Then oValue = 10000
    If oValue <> lMemorySize Then NeedSave = True
    lMemorySize = oValue
End Property
