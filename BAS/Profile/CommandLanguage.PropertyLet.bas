Property Let CommandLanguage(oValue As CommandLanguages)
    If oValue <> lCommandLanguage Then NeedSave = True
    lCommandLanguage = oValue
End Property
