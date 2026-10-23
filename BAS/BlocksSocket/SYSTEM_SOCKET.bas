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
