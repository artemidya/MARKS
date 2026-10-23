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
