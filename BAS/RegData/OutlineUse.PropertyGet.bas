Property Get OutlineUse() As Boolean
    If GetSetting(AppName, "Duplicates", "OutlineUse") = "1" Then OutlineUse = True
End Property
