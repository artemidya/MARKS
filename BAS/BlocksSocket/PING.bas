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
