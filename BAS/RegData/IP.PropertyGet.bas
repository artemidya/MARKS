Property Get IP(ProfileNumber As String) As String
    IP = Trim(GetSetting(AppName, "Profile" & ProfileNumber, "IP"))
    If IP = "" Then
        IP = "192.168.1.1"
    End If
End Property
