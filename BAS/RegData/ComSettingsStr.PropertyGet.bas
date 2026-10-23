Property Get ComSettingsStr(ComPortNumber As String) As String
    On Error GoTo 10
     ComSettingsStr = Trim(GetSetting(AppName, "COM", UCase(Trim(ComPortNumber))))
     If ComSettingsStr = "" Then
        ComSettingsStr = "baud=9600 parity=N data=8 stop=1 xon=off odsr=off dtr=on rts=on"
     End If
     Exit Property
10
    ComSettingsStr = "baud=9600 parity=N data=8 stop=1 xon=off odsr=off dtr=on rts=on"
End Property
