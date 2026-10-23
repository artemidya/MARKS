Public Function StartWinsock(Optional sDescription) As Boolean
    If IsHooked = False Then MsgBox "No Hooked Form! Error": Exit Function
    CountCreationSockets = CountCreationSockets + 1
    If Not WSAStartedUp Then
        Dim StartupData As WSADataType
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        End If
    End If
    StartWinsock = WSAStartedUp
End Function
