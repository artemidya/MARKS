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
