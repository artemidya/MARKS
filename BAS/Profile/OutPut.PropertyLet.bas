Public Property Let OutPut(oValue As String)
    If oValue <> lOutPut Then NeedSave = True
    lOutPut = oValue
End Property
