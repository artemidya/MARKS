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
