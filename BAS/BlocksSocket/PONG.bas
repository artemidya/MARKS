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
