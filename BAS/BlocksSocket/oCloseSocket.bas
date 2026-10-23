Private Sub oCloseSocket()
    'Override closesocket
    On Error Resume Next
    Dim skp As Boolean
    'Skip
    Select Case SocketStatus
        Case WaitingStarter
            If RemainderCount > 0 Then skp = True
            SkipBytes 0
        Case WaitingLengthDatagramm
            SkipBytes 4
            skp = True
        Case WaitingZamTerminator
            SkipBytes 8
            skp = True
        Case Reading
            SkipBytes CP_Datagramm + 12
            skp = True
    End Select
    
    LengthDatagramm = 0
    CP_Datagramm = 0
    zProtected_RaiseEvent_OnStatisticChange Me
    If skp Then zProtected_RaiseEvent_OnSkip Me
    closesocket (lSocket)
    BlocksSocketShared.OpenedSockets.Remove CStr(lSocket)
    If zProtected_MeHaweTheParentFlag Then Parent.Clients.Remove CStr(lSocket)
    zProtected_Connected = False
    zProtected_RaiseEvent_OnDisconnect Me
    lSocket = -1
    Statistic_PING = -1
    
    
End Sub
