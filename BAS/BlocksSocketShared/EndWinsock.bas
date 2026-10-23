Public Sub EndWinsock()
    CountCreationSockets = CountCreationSockets - 1
    If CountCreationSockets > 0 Then Exit Sub
    If CountCreationSockets < 0 Then CountCreationSockets = 0

    Dim ret As Long
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
    UnHookForm
End Sub
