Public Property Let FilePLT(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lFilePLT Then NeedSave = True
    lFilePLT = oValue
End Property
