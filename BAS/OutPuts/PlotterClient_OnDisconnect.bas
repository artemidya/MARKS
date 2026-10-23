Private Sub PlotterClient_OnDisconnect(client As BlocksSocket)
    PlotterClientConnected = False
    PlotterClientWaitingStatus = False
    PlotterClientlastOperationTime = Timer
    PlotterClientInitialized = False
    'R E
    RaiseEvent WiFiChange
End Sub
