Property Get Sensitivity() As Long
    On Error GoTo 10
    Sensitivity = CLng(GetSetting(AppName, "Duplicates", "Sensitivity"))
    
10:
    If Sensitivity < 25 Then Sensitivity = 25
    If Sensitivity > 10000 Then Sensitivity = 10000
End Property
