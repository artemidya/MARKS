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
