Private Sub Timer1_OnTimer()
    If Connected Then
        With BlocksSocket1
            If .Statistic_PING < 1000 Then
                If .Statistic_PING = LastShowedPing Then .Statistic_PING = .Statistic_PING - 1
                Label11.Caption = "Ping " & .Statistic_PING & " ms"
                LastShowedPing = .Statistic_PING
                .PING
                .Statistic_PING = .Statistic_PING + 400
            Else
                Label11.Caption = ""
                .PING
            End If
        End With
    Else
        Label11.Caption = ""
    End If
End Sub
