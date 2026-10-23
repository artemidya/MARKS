Property Get AutoUpdate() As Boolean
    If GetSetting(AppName, "Settings", "AutoUpdate") = "1" Then AutoUpdate = True
End Property
