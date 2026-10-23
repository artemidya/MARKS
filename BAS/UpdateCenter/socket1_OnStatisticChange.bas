Private Sub socket1_OnStatisticChange(client As BlocksSocket)
    If UpdateProcessingStatus = DownloadingPlotCalcGMS Then
        tt = 0
        RaiseEvent DownloadPercent(client.Statistic_PercentReceiving)
    End If
End Sub
