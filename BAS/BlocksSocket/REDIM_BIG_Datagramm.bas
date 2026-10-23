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
