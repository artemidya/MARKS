Property Let name(oValue As String)
    If Trim(oValue) <> Trim(lName) Then NeedSave = True
    lName = Trim(oValue)
End Property
