Public Sub Connect(ByVal addr As String, ByVal Port As Integer)
    If lSocket > 0 Then Exit Sub
    Statistic_BytesSent = 0
    Statistic_BytesReceived = 0
    Statistic_BytesReceivedOld = 0
    Statistic_BytesSkipped = 0
    Statistic_PING = -1
    Erase SndQueueElements
    LengthDatagramm = 0
    CP_Datagramm = 0
    SocketStatus = WaitingStarter
    Address = addr & ":" & Port
    RemainderCount = 0
    RaiseEvent OnConnectionProcessIsStarted
    zProtected_RaiseEvent_OnStatisticChange Me
    ConnectSock addr, Port
End Sub
