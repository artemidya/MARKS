Property Let IP(ProfileNumber As String, oValue As String)
    On Error GoTo 10
    SaveSetting AppName, "Profile" & ProfileNumber, "IP", Trim(oValue)
10
End Property
