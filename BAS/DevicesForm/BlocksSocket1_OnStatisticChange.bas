Private Sub BlocksSocket1_OnStatisticChange(client As BlocksSocket)
    If LoadPlotterProcess = True Then
        Label40.Caption = client.Statistic_PercentReceiving & "%"
        DoEvents
    End If
End Sub
