Public Property Let ForModeratorDevice(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lForModeratorDevice Then NeedSave = True
    lForModeratorDevice = oValue
End Property
