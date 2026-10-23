Property Get Port(ProfileNumber As String) As Integer
    On Error GoTo 10
    Port = CInt(Trim(GetSetting(AppName, "Profile" & ProfileNumber, "Port")))
    

10
    If Port = 0 Then
        Port = 9100
    End If
End Property
