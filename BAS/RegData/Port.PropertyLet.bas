Property Let Port(ProfileNumber As String, oValue As Integer)
    On Error GoTo 10
    SaveSetting AppName, "Profile" & ProfileNumber, "Port", oValue
10
End Property
