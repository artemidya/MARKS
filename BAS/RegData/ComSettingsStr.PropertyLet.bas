Property Let ComSettingsStr(ComPortNumber As String, oValue As String)
    On Error GoTo 10
    SaveSetting AppName, "COM", UCase(Trim(ComPortNumber)), oValue
10
End Property
