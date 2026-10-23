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
