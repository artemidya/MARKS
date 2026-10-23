Private Sub ConnectSock(ByVal host As String, ByVal Port As Integer, Optional ByVal IsBlocking As Boolean = False)
    Dim sockin As sockaddr
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then Exit Sub
    sockin.sin_addr = GetHostByNameAlias(host$)
    If sockin.sin_addr = INADDR_NONE Then Exit Sub
    lSocket = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If lSocket < 0 Then Exit Sub
    BlocksSocketShared.OpenedSockets.Add Me, CStr(lSocket)
    If SetSockLinger(lSocket, 0, 0) = SOCKET_ERROR Then oCloseSocket: Exit Sub
    If IsBlocking = True Then
        '//For Blocking Socket we call WSAAsyncSelect after connect
        If WSConnect(lSocket, sockin, sockaddr_size) <> 0 Then oCloseSocket: Exit Sub
        If WSAAsyncSelect(lSocket, Hooked_hWnd, ByVal WINSOCKMSG, FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE) Then oCloseSocket: Exit Sub
    Else
        '//For Non-Blocking Socket we call WSAAsyncSelect before connect
        If WSAAsyncSelect(lSocket, Hooked_hWnd, ByVal WINSOCKMSG, FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE) Then oCloseSocket: Exit Sub
        If WSConnect(lSocket, sockin, sockaddr_size) <> -1 Then oCloseSocket: Exit Sub
    End If
End Sub
