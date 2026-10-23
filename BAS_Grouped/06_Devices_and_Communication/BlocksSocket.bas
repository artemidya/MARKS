' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: BlocksSocket.cls
' Procedures: 47
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 219-224
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    StartWinsock
    RcvFlags.Denominator_2bit = DefaultDenominator
    SndFlags.Denominator_2bit = DefaultDenominator
    Statistic_PING = -1
End Sub

' -----------------------------------------------------------------------------
' DeliveryReport (PropertyGet), source lines 225-227
' -----------------------------------------------------------------------------
Property Get DeliveryReport() As Boolean
    DeliveryReport = lWaitingDeliveryReport
End Property

' -----------------------------------------------------------------------------
' DeliveryReport (PropertyLet), source lines 228-248
' -----------------------------------------------------------------------------
Property Let DeliveryReport(value As Boolean)
    lWaitingDeliveryReport = value
    'Send about NEED DELIVERY
    If RcvFlags.Denominator_2bit < SndFlags.Denominator_2bit Then SndFlags.Denominator_2bit = RcvFlags.Denominator_2bit
    SndFlags.SystemFlag = 1
    Dim b(3) As Byte
    If value Then
        MemMove b(0), htonl(SYSTEM_SOCKET_COMMANDS.sSET_DELIVERY_REPORT_ON), 4
    Else
        MemMove b(0), htonl(SYSTEM_SOCKET_COMMANDS.sSET_DELIVERY_REPORT_OFF), 4
    End If
    Dim lngBuffer() As Long
    lngBuffer = CreateLongPacket(b, 4)
    If Not Not SndQueueElements Then
        AddToSndQueueLng lngBuffer
    Else
        If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
            AddToSndQueueLng lngBuffer
        End If
    End If
End Property

' -----------------------------------------------------------------------------
' Connected (PropertyGet), source lines 249-251
' -----------------------------------------------------------------------------
Property Get Connected() As Boolean
    Connected = zProtected_Connected
End Property

' -----------------------------------------------------------------------------
' Statistic_BytesReceiving (PropertyGet), source lines 252-258
' -----------------------------------------------------------------------------
Public Property Get Statistic_BytesReceiving() As Long
    If LengthBIGDatagramm > 0 Then
        Statistic_BytesReceiving = CP_BIGDatagramm + CP_Datagramm
    Else
        Statistic_BytesReceiving = CP_Datagramm
    End If
End Property

' -----------------------------------------------------------------------------
' Statistic_BytesReceivingSize (PropertyGet), source lines 259-265
' -----------------------------------------------------------------------------
Public Property Get Statistic_BytesReceivingSize() As Long
    If LengthBIGDatagramm > 0 Then
        Statistic_BytesReceivingSize = LengthBIGDatagramm
    Else
        Statistic_BytesReceivingSize = LengthDatagramm
    End If
End Property

' -----------------------------------------------------------------------------
' Statistic_PercentReceiving (PropertyGet), source lines 266-311
' -----------------------------------------------------------------------------
Public Property Get Statistic_PercentReceiving() As Long
    If Not Clients Is Nothing Then
        Static Percent As Long
        Static client As BlocksSocket
        For Each client In Clients
            With client
                If .LengthDatagramm > 0 Or .LengthBIGDatagramm > 0 Then
                Percent = .Statistic_PercentReceiving
                If Percent > Statistic_PercentReceiving Then Statistic_PercentReceiving = Percent
                End If
            End With
        Next
        Exit Property
    End If
    Static lth As Long
    Static dtg As Long
    If LengthBIGDatagramm > 0 Then
        If LengthBIGDatagramm >= 100 Then
            lth = LengthBIGDatagramm \ 100
            dtg = CP_BIGDatagramm - LengthBIGDatagramm Mod 100
            Statistic_PercentReceiving = dtg \ lth
        ElseIf LengthBIGDatagramm >= 10 Then
            lth = LengthBIGDatagramm \ 10
            dtg = CP_BIGDatagramm - LengthBIGDatagramm Mod 10
            Statistic_PercentReceiving = 10 * dtg \ lth
        Else
            lth = LengthBIGDatagramm
            dtg = CP_BIGDatagramm
            Statistic_PercentReceiving = 100 * dtg \ lth
        End If
    Else
        If LengthDatagramm >= 100 Then
            lth = LengthDatagramm \ 100
            dtg = CP_Datagramm - LengthDatagramm Mod 100
            Statistic_PercentReceiving = dtg \ lth
        ElseIf LengthDatagramm >= 10 Then
            lth = LengthDatagramm \ 10
            dtg = CP_Datagramm - LengthDatagramm Mod 10
            Statistic_PercentReceiving = 10 * dtg \ lth
        ElseIf LengthDatagramm > 0 Then
            lth = LengthDatagramm
            dtg = CP_Datagramm
            Statistic_PercentReceiving = 100 * dtg \ lth
        End If
    End If
End Property

' -----------------------------------------------------------------------------
' Statistic_All (PropertyGet), source lines 312-368
' -----------------------------------------------------------------------------
Public Property Get Statistic_All(Optional ShowRcvFlags As Boolean = False) As String
    Dim q As Long, Percent As Long: Percent = Statistic_PercentReceiving
    Dim PBar As String ' [+++______________]
    Dim oldRcv As String
    PBar = "["
    For q = 1 To Percent \ 5
        PBar = PBar & "+"
    Next
    For q = q To 20
        PBar = PBar & "_"
    Next
    PBar = PBar & "] "
    If OnReciveOld_TurnON Then oldRcv = "; ReceivedOld=" & Statistic_BytesReceivedOld
    Dim sum As String
    If LengthBIGDatagramm > 0 Then
        sum = CP_BIGDatagramm & "/" & LengthBIGDatagramm
    Else
        sum = CP_Datagramm & "/" & LengthDatagramm
    End If
    Statistic_All = Address & " Sent=" & Statistic_BytesSent _
                            & "; Received=" & Statistic_BytesReceived _
                            & oldRcv _
                            & "; Skipped=" & Statistic_BytesSkipped _
                            & "; ProgressReceiving=" & PBar & Percent & "% " _
                            & sum _
                            & "; PING=" & Statistic_PING & " ms"
    If Not (Clients Is Nothing) Then
        Statistic_All = Statistic_All & vbCrLf & "==================================================" & vbCrLf
        Statistic_All = Statistic_All & "Clients.Count=" & Clients.count & vbCrLf
        Statistic_All = Statistic_All & "==================================================" & vbCrLf
        Dim client As BlocksSocket
        For Each client In Clients
             Statistic_All = Statistic_All & client.Statistic_All(ShowRcvFlags) & vbCrLf
        Next
    Else
        If ShowRcvFlags Then
            With RcvFlags
                Statistic_All = Statistic_All & vbCrLf & "RcvFlags> NeedReplaceFlag=" & .NeedReplaceFlag _
                                                       & "; CompressFlag=" & .CompressFlag _
                                                       & "; CryptoFlag=" & .CryptoFlag _
                                                       & "; PartFlag=" & .PartFlag _
                                                       & "; SendFilesFlag=" & .SendFilesFlag _
                                                       & "; SystemFlag=" & .SendFilesFlag _
                                                       & "; Denominator=" & .Denominator_2bit
            End With
            With SndFlags
                Statistic_All = Statistic_All & vbCrLf & "SndFlags> NeedReplaceFlag=" & .NeedReplaceFlag _
                                                       & "; CompressFlag=" & .CompressFlag _
                                                       & "; CryptoFlag=" & .CryptoFlag _
                                                       & "; PartFlag=" & .PartFlag _
                                                       & "; SendFilesFlag=" & .SendFilesFlag _
                                                       & "; SystemFlag=" & .SendFilesFlag _
                                                       & "; Denominator=" & .Denominator_2bit & vbCrLf
            End With
        End If
    End If
End Property

' -----------------------------------------------------------------------------
' Connect (Sub), source lines 369-385
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' ConnectSock (Sub), source lines 386-406
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' SetSockLinger (Function), source lines 407-420
' -----------------------------------------------------------------------------
Private Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
    Dim linger As LingerType
    linger.l_onoff = OnOff
    linger.l_linger = LingerTime
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, linger, 4) Then
        'Debug.Print "Error setting linger info: " & WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
    Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, linger, 4) Then
            'Debug.Print "Error getting linger info: " & WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' GetHostByNameAlias (Function), source lines 421-440
' -----------------------------------------------------------------------------
Private Function GetHostByNameAlias(ByVal hostname As String) As Long
    On Error Resume Next
    Dim phe As LongPtr
    Dim heDestHost As HostEnt
    Dim addrList As LongPtr
    Dim retIP As Long
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            MemMove heDestHost, ByVal phe, LenB(heDestHost)
            MemMove addrList, ByVal heDestHost.h_addr_list, LenB(heDestHost.h_addr_list)
            MemMove retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function

' -----------------------------------------------------------------------------
' getascip (Function), source lines 441-458
' -----------------------------------------------------------------------------
Private Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemMove ByVal retString, ByVal lpStr, nStr
    retString = left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function

' -----------------------------------------------------------------------------
' Send (Sub), source lines 459-510
' -----------------------------------------------------------------------------
Public Sub Send(ByRef bytes() As Byte)
    If Not Not bytes Then
        'Statistic
        Dim lngBuffer() As Long
        If RcvFlags.Denominator_2bit < SndFlags.Denominator_2bit Then SndFlags.Denominator_2bit = RcvFlags.Denominator_2bit
        If UBound(bytes) + 1 > SndFlags.Denominator_2bit Then
            'System, need sent BigLengthDatagramm
            REDIM_BIG_Datagramm UBound(bytes) + 1
            SndFlags.SystemFlag = 0
            SndFlags.PartFlag = 1
            Dim sp As Long
            If Not Not SndQueueElements Then
                AddToSndQueueBytes bytes, sp, 1, FlagsToByte(SndFlags)
            Else
                Dim ostatok As Long: ostatok = (UBound(bytes) + 1) Mod SndFlags.Denominator_2bit
                Dim q As Long
                For q = 1 To (UBound(bytes) + 1) \ SndFlags.Denominator_2bit
                    lngBuffer = CreateLongPacket(bytes, SndFlags.Denominator_2bit, sp)
                    sp = sp + SndFlags.Denominator_2bit
                    If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                        AddToSndQueueLng lngBuffer
                        AddToSndQueueBytes bytes, sp, 1, FlagsToByte(SndFlags)
                        Exit Sub
                    Else
                        AddToStatisticSent (UBound(lngBuffer) + 1) * 4 - 12
                    End If
                Next
                If ostatok > 0 Then
                    lngBuffer = CreateLongPacket(bytes, ostatok, sp)
                    If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                        AddToSndQueueLng lngBuffer
                    Else
                        AddToStatisticSent (UBound(lngBuffer) + 1) * 4 - 12
                    End If
                End If
            End If
        Else
            SndFlags.SystemFlag = 0
            SndFlags.PartFlag = 0
            lngBuffer = CreateLongPacket(bytes, UBound(bytes) + 1)
            If Not Not SndQueueElements Then
                AddToSndQueueLng lngBuffer
            Else
                If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                    AddToSndQueueLng lngBuffer
                Else
                    AddToStatisticSent (UBound(lngBuffer) + 1) * 4 - 12
                End If
            End If
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' SendOld (Sub), source lines 511-524
' -----------------------------------------------------------------------------
Public Sub SendOld(ByRef bytes() As Byte)
    If Not Not bytes Then
        If Not Not SndQueueElements Then
            AddToSndQueueBytes bytes()
        Else
            If SendLng(lSocket, bytes(0), UBound(bytes) + 1, 0) < 1 Then
                AddToSndQueueBytes bytes()
            Else
                'Add Statistic
                AddToStatisticSent UBound(bytes) + 1
            End If
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' AddToStatisticSent (Sub), source lines 525-534
' -----------------------------------------------------------------------------
Private Sub AddToStatisticSent(ByVal cntBytes As Long)
    'Inline
    Statistic_BytesSent = Statistic_BytesSent + cntBytes
    If zProtected_MeHaweTheParentFlag Then
        With Parent
            .Statistic_BytesSent = .Statistic_BytesSent + cntBytes
        End With
    End If
    zProtected_RaiseEvent_OnStatisticChange Me
End Sub

' -----------------------------------------------------------------------------
' AddToSndQueueLng (Sub), source lines 535-542
' -----------------------------------------------------------------------------
Private Sub AddToSndQueueLng(ByRef lngBuffer() As Long)
    If Not Not SndQueueElements Then
        ReDim Preserve SndQueueElements(UBound(SndQueueElements) + 1)
    Else
        ReDim SndQueueElements(0)
    End If
    SndQueueElements(UBound(SndQueueElements)).LngBuff = lngBuffer
End Sub

' -----------------------------------------------------------------------------
' AddToSndQueueBytes (Sub), source lines 543-555
' -----------------------------------------------------------------------------
Private Sub AddToSndQueueBytes(ByRef bytes() As Byte, Optional ByVal cp As Long = 0, Optional ByVal needConvertToLng As Long = 0, Optional ByVal Flags As Byte = 0)
    If Not Not SndQueueElements Then
        ReDim Preserve SndQueueElements(UBound(SndQueueElements) + 1)
    Else
        ReDim SndQueueElements(0)
    End If
    With SndQueueElements(UBound(SndQueueElements))
        .ByteBuff() = bytes()
        .cp = cp
        .Flags = Flags
        .needConvertToLng = needConvertToLng
    End With
End Sub

' -----------------------------------------------------------------------------
' ProcessMessage (Sub), source lines 556-648
' -----------------------------------------------------------------------------
Public Sub ProcessMessage(ByVal mess As Long) '======================================================================================
    Select Case mess
        Case FD_ACCEPT      'new connection for Server only
            Dim lngRetValue As Long, udtSocketAddress As sockaddr
            lngRetValue = accept(lSocket, udtSocketAddress, LenB(udtSocketAddress))
            If lngRetValue <> INVALID_SOCKET Then
                Dim newClient As New BlocksSocket
                With newClient
                    .zProtected_MeHaweTheParentFlag = 1
                    Set .Parent = Me
                    .lSocket = lngRetValue
                    .Address = getascip(udtSocketAddress.sin_addr) & ":" & 65536 + ntohs(udtSocketAddress.sin_port)
                    .OnReciveOld_TurnON = Me.OnReciveOld_TurnON
                    BlocksSocketShared.OpenedSockets.Add newClient, CStr(.lSocket)
                    Clients.Add newClient, CStr(.lSocket)
                    newClient.zProtected_Connected = True
                    zProtected_RaiseEvent_OnConnect newClient
                End With
            End If
            Exit Sub
        Case FD_READ        '1 - we have data waiting to be processed
            RecivL
            Exit Sub
        Case FD_CONNECT      '16 - we are connected to server
            zProtected_Connected = True
            zProtected_RaiseEvent_OnConnect Me
            Exit Sub
        Case FD_WRITE       '2  - we can write to our connection
            Dim needDel As Long
            Dim toStatistic As Long
            Do While Not Not SndQueueElements()
                'need sndQueueProcessing ???
                With SndQueueElements(0)
                    If Not Not .LngBuff() Then
                        If SendLng(lSocket, .LngBuff(0), (UBound(.LngBuff()) + 1) * 4, 0) < 1 Then
                            Exit Sub
                        Else
                            needDel = 1
                            toStatistic = (UBound(.LngBuff()) + 1) * 4 - 12
                        End If
                    Else
                        If .needConvertToLng Then
                            'send packet
                            SndFlags = ByteToFlags(.Flags)
                            Dim q As Long
                            Dim lngBuffer() As Long
                            For q = 1 To (UBound(.ByteBuff) + 1 - .cp) \ SndFlags.Denominator_2bit
                                'createpacket
                                lngBuffer = CreateLongPacket(.ByteBuff, SndFlags.Denominator_2bit, .cp)
                                If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                                    AddToStatisticSent toStatistic
                                    Exit Sub
                                Else
                                    toStatistic = toStatistic + (UBound(lngBuffer) + 1) * 4 - 12
                                End If
                                .cp = .cp + SndFlags.Denominator_2bit
                            Next
                            Dim ostatok As Long: ostatok = (UBound(.ByteBuff) + 1) Mod SndFlags.Denominator_2bit
                            If ostatok > 0 Then
                                lngBuffer = CreateLongPacket(.ByteBuff, ostatok, .cp)
                                If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                                    Exit Sub
                                Else
                                    toStatistic = toStatistic + (UBound(lngBuffer) + 1) * 4 - 12
                                End If
                            End If
                            needDel = 1
                        Else
                            'send old
                            If SendLng(lSocket, .ByteBuff(0), UBound(.ByteBuff()) + 1, 0) < 1 Then
                                Exit Sub
                            Else
                                needDel = 1
                                toStatistic = UBound(.ByteBuff) + 1
                            End If
                        End If
                    End If
                End With
                If needDel Then
                    DelZerroSndQueueElement
                    needDel = 0
                    AddToStatisticSent toStatistic
                    toStatistic = 0
                End If
            Loop
            Exit Sub
        Case FD_CLOSE       '32 - the connection with server is closed
            RecivL
            Exit Sub
    End Select
    'Errore Need close socket
    oCloseSocket
End Sub

' -----------------------------------------------------------------------------
' DelZerroSndQueueElement (Sub), source lines 649-659
' -----------------------------------------------------------------------------
Private Sub DelZerroSndQueueElement()
    If UBound(SndQueueElements()) = 0 Then
        Erase SndQueueElements()
    Else
        Dim q As Long
        For q = 0 To UBound(SndQueueElements()) - 1
            SndQueueElements(q) = SndQueueElements(q + 1)
        Next
        ReDim Preserve SndQueueElements(UBound(SndQueueElements()) - 1)
    End If
End Sub

' -----------------------------------------------------------------------------
' RecivL (Sub), source lines 660-810
' -----------------------------------------------------------------------------
Private Sub RecivL()
    'With unpacketing
    Dim x As Long
    x = Recv(lSocket, LngReadBuffer(0), BuffSize, 0)
    If x > 0 Then
        'For old
        If OnReciveOld_TurnON = True Then
           'OnReciveOld
           Dim buffer() As Byte
           ReDim buffer(x - 1)
           MemMove buffer(0), LngReadBuffer(0), x
           Statistic_BytesReceivedOld = Statistic_BytesReceivedOld + x
           If zProtected_MeHaweTheParentFlag Then
                With Parent
                    .Statistic_BytesReceivedOld = .Statistic_BytesReceivedOld + x
                End With
           End If
           zProtected_RaiseEvent_OnStatisticChange Me
           zProtected_RaiseEvent_OnReciveOld buffer, Me
           Exit Sub
        End If

        'Remainder=================
        If RemainderCount > 0 Then
            AddArray RemainderCount, x
            x = x + RemainderCount
        End If
        
        RemainderCount = x Mod 4
        
        If RemainderCount > 0 Then
            x = x - RemainderCount
            MemMove Remainder(0), LngReadBuffer(x \ 4), 4
        End If
        
        If x < 4 Then Exit Sub
        '===========================

        Dim q As Long, lngX As Long
        lngX = (x - 1) \ 4   '<< - Ubound
        For q = 0 To lngX
            Select Case SocketStatus
                '==================================================
                Case WaitingStarter
                    If LngReadBuffer(q) = 0 Then
                        SocketStatus = WaitingLengthDatagramm
                    Else
                        Dim bytesSkipped As Long: bytesSkipped = 4
                        Do While q < lngX
                            q = q + 1
                            If LngReadBuffer(q) = 0 Then
                                SocketStatus = WaitingLengthDatagramm
                                Exit Do
                            Else
                                bytesSkipped = bytesSkipped + 4
                            End If
                        Loop
                        SkipBytes bytesSkipped
                        zProtected_RaiseEvent_OnStatisticChange Me
                        zProtected_RaiseEvent_OnSkip Me
                    End If
                '==================================================
                Case WaitingLengthDatagramm
                    If LngReadBuffer(q) = 0 Then
                        SkipBytes 4
                        zProtected_RaiseEvent_OnStatisticChange Me
                        zProtected_RaiseEvent_OnSkip Me
                    Else
                        Static b As Byte
                        MemMove b, LngReadBuffer(q), 1
                        RcvFlags = ByteToFlags(b)
                        MemMove LngReadBuffer(q), 0, 1
                        LengthDatagramm = ntohl(LngReadBuffer(q))
                        If LengthDatagramm < 1 Or LengthDatagramm > RcvFlags.Denominator_2bit Or LengthDatagramm > LimitPacket Then
                            LengthDatagramm = 0
                            SkipBytes 8
                            zProtected_RaiseEvent_OnStatisticChange Me
                            zProtected_RaiseEvent_OnSkip Me
                            SocketStatus = WaitingStarter
                        Else
                            SocketStatus = WaitingZamTerminator
                        End If
                    End If
                '==================================================
                Case WaitingZamTerminator
                    If LngReadBuffer(q) = 0 Then
                        SkipBytes 8
                        zProtected_RaiseEvent_OnStatisticChange Me
                        zProtected_RaiseEvent_OnSkip Me
                        SocketStatus = WaitingLengthDatagramm
                    Else
                        ZamTerminator = LngReadBuffer(q)
                        CP_Datagramm = 0
                        ReDim Datagramm(LengthDatagramm - 1)
                        SocketStatus = Reading
                    End If
                '==================================================
                Case Reading
                    Dim s As Long, e As Long
                    e = (LengthDatagramm - CP_Datagramm - 1) \ 4 + q
                    If e > lngX Then e = lngX
                    'Proverka na 0
                    For s = q To e
                        If LngReadBuffer(s) = 0 Then
                            'Skip
                            SkipBytes CP_Datagramm + (s - q) * 4 + 12
                            zProtected_RaiseEvent_OnStatisticChange Me
                            zProtected_RaiseEvent_OnSkip Me
                            SocketStatus = WaitingLengthDatagramm
                            q = s
                            GoTo nx
                        End If
                    Next
                    If RcvFlags.NeedReplaceFlag Then
                        'Zamena terminatora
                        For s = q To e
                            If LngReadBuffer(s) = ZamTerminator Then LngReadBuffer(s) = 0
                        Next
                    End If
                    'Copy to Datagramm result
                    s = (e - q + 1) * 4
                    If s > LengthDatagramm - CP_Datagramm Then s = LengthDatagramm - CP_Datagramm
                    MemMove Datagramm(CP_Datagramm), LngReadBuffer(q), s
                    CP_Datagramm = CP_Datagramm + s
                    
                    If GetTickCount - LastRefresTiks > 200 And RcvFlags.SystemFlag = 0 Then
                        If RcvFlags.SystemFlag = 0 Then
                            If DeliveryReportRequired Then
                                If GetTickCount - LastSendDeliveryReportTiks > 1000 Then SendDeliveryReport
                            End If
                            zProtected_RaiseEvent_OnStatisticChange Me
                        End If
                    End If
                    If CP_Datagramm >= LengthDatagramm Then
                        'All Datagramm Received
                        SocketStatus = WaitingStarter
                        '1. Decript Datagramm()
                        '2. UnZIP Datagramm()
                        '3
                        AllDatagrammReceived
                    End If
                    q = e
                '==================================================
            End Select
nx:
        Next
    Else
        shutdown lSocket, SD_SEND
        oCloseSocket
    End If
End Sub

' -----------------------------------------------------------------------------
' AddArray (Sub), source lines 811-818
' -----------------------------------------------------------------------------
Private Sub AddArray(length1 As Long, length2 As Long)
    Dim ms() As Byte
    ReDim ms(length1 + length2 - 1)
    
    MemMove ms(0), Remainder(0), length1
    MemMove ms(length1), LngReadBuffer(0), length2
    MemMove LngReadBuffer(0), ms(0), length1 + length2
End Sub

' -----------------------------------------------------------------------------
' SkipBytes (Sub), source lines 819-836
' -----------------------------------------------------------------------------
Private Static Sub SkipBytes(ByVal BytesCount As Long)
    'inline
    BytesCount = BytesCount + RemainderCount
    RemainderCount = 0
    If LengthBIGDatagramm > 0 Then
        'SKIP BIG
        BytesCount = BytesCount + CP_BIGDatagramm
        LengthBIGDatagramm = 0
        CP_BIGDatagramm = 0
        Erase BIGDatagramm
    End If
    Statistic_BytesSkipped = Statistic_BytesSkipped + BytesCount
    If zProtected_MeHaweTheParentFlag Then
        With Parent
            .Statistic_BytesSkipped = .Statistic_BytesSkipped + BytesCount
        End With
    End If
End Sub

' -----------------------------------------------------------------------------
' AllDatagrammReceived (Sub), source lines 837-882
' -----------------------------------------------------------------------------
Private Sub AllDatagrammReceived()
    If RcvFlags.SystemFlag Then
        LengthDatagramm = 0
        CP_Datagramm = 0
        SYSTEM_SOCKET Datagramm()
    Else
        If RcvFlags.PartFlag Then
            MemMove BIGDatagramm(CP_BIGDatagramm), Datagramm(0), UBound(Datagramm()) + 1
            CP_BIGDatagramm = CP_BIGDatagramm + UBound(Datagramm()) + 1
            'LengthBIGDatagramm
            If CP_BIGDatagramm = LengthBIGDatagramm Then
                Statistic_BytesReceived = Statistic_BytesReceived + LengthBIGDatagramm
                If zProtected_MeHaweTheParentFlag Then
                    With Parent
                        .Statistic_BytesReceived = .Statistic_BytesReceived + LengthBIGDatagramm
                    End With
                End If
                LengthDatagramm = 0
                CP_Datagramm = 0
                If DeliveryReportRequired Then SendDeliveryReport
                zProtected_RaiseEvent_OnStatisticChange Me
                zProtected_RaiseEvent_OnRecive BIGDatagramm(), Me
                LengthBIGDatagramm = 0
                CP_BIGDatagramm = 0
                zProtected_RaiseEvent_OnStatisticChange Me
            Else
                If GetTickCount - LastRefresTiks > 200 Then
                    zProtected_RaiseEvent_OnStatisticChange Me
                End If
            End If
        Else
            Statistic_BytesReceived = Statistic_BytesReceived + LengthDatagramm
            If zProtected_MeHaweTheParentFlag Then
                With Parent
                    .Statistic_BytesReceived = .Statistic_BytesReceived + LengthDatagramm
                End With
            End If
            If DeliveryReportRequired Then SendDeliveryReport
            zProtected_RaiseEvent_OnStatisticChange Me
            zProtected_RaiseEvent_OnRecive Datagramm(), Me
            LengthDatagramm = 0
            CP_Datagramm = 0
            zProtected_RaiseEvent_OnStatisticChange Me
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' oCloseSocket (Sub), source lines 883-916
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' Disconnect (Sub), source lines 917-925
' -----------------------------------------------------------------------------
Public Sub Disconnect()
'    If shutdown(lSocket, SD_BOTH) < 0 Then
'        oCloseSocket
'    End If
    
    shutdown lSocket, SD_SEND
    
    oCloseSocket
End Sub

' -----------------------------------------------------------------------------
' CreateLongPacket (Function), source lines 942-966
' -----------------------------------------------------------------------------
Public Function CreateLongPacket(ByRef ByteBuffer() As Byte, ByVal lengthByteBuffer As Long, Optional ByVal sp As Long = 0) As Long()    'Do 2 Gb
    Dim lengthLongBuffer As Long
    lengthLongBuffer = (lengthByteBuffer + 3) \ 4 + 3
    Dim longBuffer() As Long
    ReDim longBuffer(lengthLongBuffer - 1)
    MemMove longBuffer(3), ByteBuffer(sp), lengthByteBuffer
    longBuffer(1) = htonl(lengthByteBuffer)
    Dim ZamTerminator As Long
    'ZamTerminator = 0
    Do While FindLong(ZamTerminator, longBuffer)
        ZamTerminator = NewZamTerminator
    Loop
    If ZamTerminator = 0 Then
        'No replace
        ZamTerminator = NewZamTerminator
        SndFlags.NeedReplaceFlag = 0
    Else
        'Replace
        Replace_0_ToZamTerminatorLong ZamTerminator, longBuffer
        SndFlags.NeedReplaceFlag = 1
    End If
    MemMove longBuffer(1), FlagsToByte(SndFlags), 1
    longBuffer(2) = ZamTerminator
    CreateLongPacket = longBuffer
End Function

' -----------------------------------------------------------------------------
' FindLong (Function), source lines 967-975
' -----------------------------------------------------------------------------
Private Function FindLong(ByVal value As Long, ByRef longBuffer() As Long) As Boolean
    Dim q As Long
    For q = 3 To UBound(longBuffer)
        If value = longBuffer(q) Then
            FindLong = True
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------
' Replace_0_ToZamTerminatorLong (Sub), source lines 976-981
' -----------------------------------------------------------------------------
Private Sub Replace_0_ToZamTerminatorLong(ByVal value As Long, ByRef longBuffer() As Long)
    Dim q As Long
    For q = 3 To UBound(longBuffer)
        If 0 = longBuffer(q) Then longBuffer(q) = value
    Next
End Sub

' -----------------------------------------------------------------------------
' NewZamTerminator (Function), source lines 982-1000
' -----------------------------------------------------------------------------
Private Function NewZamTerminator() As Long
    'Delaet noviy zamenitel dla terminatora (0.0.0.0)
1:
    On Error GoTo 10
    Static Z As Long
    If Z = 0 Then
        Randomize Timer
        Z = CLng(Rnd * 2147483647)
    End If
    If Z < 2147483647 Then
        Z = Z + 1
    Else
        Z = -2147483647 - 1
    End If
    NewZamTerminator = Z
    Exit Function
10
      GoTo 1
End Function

' -----------------------------------------------------------------------------
' FlagsToByte (Function), source lines 1001-1019
' -----------------------------------------------------------------------------
Private Static Function FlagsToByte(Flags As SocketSystemFlags) As Byte
    With Flags
        FlagsToByte = .NeedReplaceFlag * 128
        FlagsToByte = FlagsToByte Or .CompressFlag * 64
        FlagsToByte = FlagsToByte Or .CryptoFlag * 32
        FlagsToByte = FlagsToByte Or .PartFlag * 16
        FlagsToByte = FlagsToByte Or .SendFilesFlag * 8
        FlagsToByte = FlagsToByte Or .SystemFlag * 4
        If .Denominator_2bit = 16777212 Then
            Exit Function
        ElseIf .Denominator_2bit = 4194304 Then
            FlagsToByte = FlagsToByte Or 1
        ElseIf .Denominator_2bit = 65536 Then
            FlagsToByte = FlagsToByte Or 2
        Else    'If .Denominator_2bit = 1024
            FlagsToByte = FlagsToByte Or 3
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' ByteToFlags (Function), source lines 1020-1038
' -----------------------------------------------------------------------------
Private Static Function ByteToFlags(ByVal b As Byte) As SocketSystemFlags
    With ByteToFlags
        .NeedReplaceFlag = b \ 128
        .CompressFlag = b \ 64 And 1
        .CryptoFlag = b \ 32 And 1
        .PartFlag = b \ 16 And 1
        .SendFilesFlag = b \ 8 And 1
        .SystemFlag = b \ 4 And 1
        If (b And 3) = 0 Then
            .Denominator_2bit = 16777212
        ElseIf (b And 3) = 1 Then
            .Denominator_2bit = 4194304
        ElseIf (b And 3) = 2 Then
            .Denominator_2bit = 65536
        Else 'If(b And 3) = 3
            .Denominator_2bit = 1024
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' ServerUP (Sub), source lines 1042-1065
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' ServerDOWN (Sub), source lines 1066-1079
' -----------------------------------------------------------------------------
Public Sub ServerDOWN()
    'disconnect all
    If Not Clients Is Nothing Then
        Dim skt As BlocksSocket
        For Each skt In Clients
            skt.Disconnect
        Next
    End If
'    If Not Clients Is Nothing Then
'        MsgBox "SOME CLIENTS NOT IS CLOSED"
'    End If
    Set Clients = Nothing
    oCloseSocket
End Sub

' -----------------------------------------------------------------------------
' zProtected_RaiseEvent_OnConnect (Sub), source lines 1080-1087
' -----------------------------------------------------------------------------
Public Sub zProtected_RaiseEvent_OnConnect(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnConnect BlocksSocket
    Else
        RaiseEvent OnConnect(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub

' -----------------------------------------------------------------------------
' zProtected_RaiseEvent_OnDisconnect (Sub), source lines 1088-1095
' -----------------------------------------------------------------------------
Public Sub zProtected_RaiseEvent_OnDisconnect(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnDisconnect BlocksSocket
    Else
        RaiseEvent OnDisconnect(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub

' -----------------------------------------------------------------------------
' zProtected_RaiseEvent_OnStatisticChange (Sub), source lines 1096-1103
' -----------------------------------------------------------------------------
Public Sub zProtected_RaiseEvent_OnStatisticChange(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnStatisticChange BlocksSocket
    Else
        RaiseEvent OnStatisticChange(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub

' -----------------------------------------------------------------------------
' zProtected_RaiseEvent_OnRecive (Sub), source lines 1104-1111
' -----------------------------------------------------------------------------
Public Sub zProtected_RaiseEvent_OnRecive(ByRef bytes() As Byte, ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnRecive bytes, BlocksSocket
    Else
        RaiseEvent OnRecive(bytes, BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub

' -----------------------------------------------------------------------------
' zProtected_RaiseEvent_OnSkip (Sub), source lines 1112-1119
' -----------------------------------------------------------------------------
Public Sub zProtected_RaiseEvent_OnSkip(ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnSkip BlocksSocket
    Else
        RaiseEvent OnSkip(BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub

' -----------------------------------------------------------------------------
' zProtected_RaiseEvent_OnReciveOld (Sub), source lines 1120-1127
' -----------------------------------------------------------------------------
Public Sub zProtected_RaiseEvent_OnReciveOld(ByRef bytes() As Byte, ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_RaiseEvent_OnReciveOld bytes, BlocksSocket
    Else
        RaiseEvent OnReciveOld(bytes, BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub

' -----------------------------------------------------------------------------
' zProtected_OnDeliveryReporting (Sub), source lines 1128-1135
' -----------------------------------------------------------------------------
Public Sub zProtected_OnDeliveryReporting(BytesDelivered As Long, BytesSize As Long, ByRef BlocksSocket As BlocksSocket)
    If zProtected_MeHaweTheParentFlag Then
        Parent.zProtected_OnDeliveryReporting BytesDelivered, BytesSize, BlocksSocket
    Else
        RaiseEvent OnDeliveryReporting(BytesDelivered, BytesSize, BlocksSocket)
        LastRefresTiks = GetTickCount
    End If
End Sub

' -----------------------------------------------------------------------------
' SYSTEM_SOCKET (Sub), source lines 1141-1180
' -----------------------------------------------------------------------------
Private Sub SYSTEM_SOCKET(ByRef bytes() As Byte)
    Dim systemCommand As SYSTEM_SOCKET_COMMANDS
    MemMove systemCommand, bytes(0), 4
    systemCommand = ntohl(systemCommand)
    Select Case systemCommand
        Case sPING  '=================================================
            If UBound(bytes()) = 7 Then
                MemMove Statistic_PING, bytes(4), 4
                Statistic_PING = ntohl(Statistic_PING)
                zProtected_RaiseEvent_OnStatisticChange Me
                PONG
            End If
        Case sPONG  '=================================================
            On Error Resume Next
            Statistic_PING = GetTickCount - LastPingTiks
            zProtected_RaiseEvent_OnStatisticChange Me
        Case sREDIM_BIG_Datagramm '===================================
            MemMove LengthBIGDatagramm, bytes(4), UBound(bytes) + 1 - 4
            LengthBIGDatagramm = ntohl(LengthBIGDatagramm)
            If LengthBIGDatagramm < 1 Or LengthBIGDatagramm > LimitPacket Then LengthBIGDatagramm = 0: Exit Sub
            CP_BIGDatagramm = 0
            ReDim BIGDatagramm(LengthBIGDatagramm - 1)
        Case sSET_DELIVERY_REPORT_ON
            LastSendDeliveryReportTiks = GetTickCount
            DeliveryReportRequired = True
        Case sSET_DELIVERY_REPORT_OFF
            DeliveryReportRequired = False
        Case sDELIVERY_REPORT
            'REPORT
             If UBound(bytes()) = 11 Then
                Dim deliveredBytes As Long
                Dim size As Long
                MemMove deliveredBytes, bytes(4), 4
                MemMove size, bytes(8), 4
                deliveredBytes = ntohl(deliveredBytes)
                size = ntohl(size)
                zProtected_OnDeliveryReporting deliveredBytes, size, Me
             End If
    End Select
End Sub

' -----------------------------------------------------------------------------
' SendDeliveryReport (Sub), source lines 1181-1203
' -----------------------------------------------------------------------------
Private Sub SendDeliveryReport()
    LastSendDeliveryReportTiks = GetTickCount
    
    Dim BytesDelivered As Long, BytesSize As Long
    BytesDelivered = Me.Statistic_BytesReceiving
    BytesSize = Me.Statistic_BytesReceivingSize
    
    SndFlags.SystemFlag = 1
    Dim b(11) As Byte
    MemMove b(0), htonl(SYSTEM_SOCKET_COMMANDS.sDELIVERY_REPORT), 4
    MemMove b(4), htonl(BytesDelivered), 4
    MemMove b(8), htonl(BytesSize), 4
    
    Dim lngBuffer() As Long
    lngBuffer = CreateLongPacket(b, UBound(b) + 1)
    If Not Not SndQueueElements Then
        AddToSndQueueLng lngBuffer
    Else
        If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
            AddToSndQueueLng lngBuffer
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' REDIM_BIG_Datagramm (Sub), source lines 1205-1219
' -----------------------------------------------------------------------------
Private Sub REDIM_BIG_Datagramm(ByVal Length As Long)   'LongLong
    SndFlags.SystemFlag = 1
    Dim b(LenB(Length) - 1 + 4) As Byte
    MemMove b(0), htonl(SYSTEM_SOCKET_COMMANDS.sREDIM_BIG_Datagramm), 4            'sREDIM_BIG_Datagramm = 3
    MemMove b(4), htonl(Length), LenB(Length)
    Dim lngBuffer() As Long
    lngBuffer = CreateLongPacket(b, UBound(b) + 1)
    If Not Not SndQueueElements Then
        AddToSndQueueLng lngBuffer
    Else
        If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
            AddToSndQueueLng lngBuffer
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' PING (Sub), source lines 1221-1244
' -----------------------------------------------------------------------------
Public Sub PING()
    If Not (Clients Is Nothing) Then
        Dim client As BlocksSocket
        For Each client In Clients
             client.PING
        Next
    Else
        If RcvFlags.Denominator_2bit < SndFlags.Denominator_2bit Then SndFlags.Denominator_2bit = RcvFlags.Denominator_2bit
        SndFlags.SystemFlag = 1
        LastPingTiks = GetTickCount
        Dim b(7) As Byte
        MemMove b(0), htonl(SYSTEM_SOCKET_COMMANDS.sPING), 4        'sPING
        MemMove b(4), htonl(Statistic_PING), 4
        Dim lngBuffer() As Long
        lngBuffer = CreateLongPacket(b, 8)
        If Not Not SndQueueElements Then
            AddToSndQueueLng lngBuffer
        Else
            If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
                AddToSndQueueLng lngBuffer
            End If
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' PONG (Sub), source lines 1245-1259
' -----------------------------------------------------------------------------
Private Sub PONG()
    If RcvFlags.Denominator_2bit < SndFlags.Denominator_2bit Then SndFlags.Denominator_2bit = RcvFlags.Denominator_2bit
    SndFlags.SystemFlag = 1
    Dim b(3) As Byte
    MemMove b(0), htonl(SYSTEM_SOCKET_COMMANDS.sPONG), 4        'sPONG
    Dim lngBuffer() As Long
    lngBuffer = CreateLongPacket(b, 4)
    If Not Not SndQueueElements Then
        AddToSndQueueLng lngBuffer
    Else
        If SendLng(lSocket, lngBuffer(0), (UBound(lngBuffer) + 1) * 4, 0) < 1 Then
            AddToSndQueueLng lngBuffer
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 1260-1266
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    If lSocket > 0 Then
        shutdown lSocket, SD_SEND
        oCloseSocket
    End If
    EndWinsock
End Sub
