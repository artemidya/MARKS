Public Sub ServerUP(ByVal localPort As Integer)
    If lSocket > 0 Then Exit Sub
    Statistic_BytesSent = 0
    Statistic_BytesReceived = 0
    Statistic_BytesReceivedOld = 0
    Statistic_BytesSkipped = 0
    LengthDatagramm = 0
    CP_Datagramm = 0
    Address = "localhost:" & localPort
    Set Clients = New Collection
    zProtected_RaiseEvent_OnStatisticChange Me
    'Call the socket Winsock API function in order create a new socket
    lSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If lSocket < 0 Then Exit Sub
    BlocksSocketShared.OpenedSockets.Add Me, CStr(lSocket)
    If WSAAsyncSelect(lSocket, BlocksSocketShared.Hooked_hWnd, WINSOCKMSG, FD_CONNECT Or FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT) Then oCloseSocket: Exit Sub
    If setsockopt(lSocket, SOL_SOCKET, SO_REUSEADDR, True, 2) Then oCloseSocket: Exit Sub
    Dim udtSocketAddress As sockaddr
    udtSocketAddress.sin_family = AF_INET
    udtSocketAddress.sin_addr = 0  ' - ALL DEVICES 'GetHostByNameAlias("0.0.0.0")
    udtSocketAddress.sin_port = htons(localPort)
    If bind(lSocket, udtSocketAddress, LenB(udtSocketAddress)) Then oCloseSocket: Exit Sub
    If listen(lSocket, SOMAXCONN) Then oCloseSocket: Exit Sub
End Sub
