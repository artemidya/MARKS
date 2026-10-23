Private Sub PlotterClient_OnConnect(client As BlocksSocket)
    PlotterClientConnected = True
    PlotterClientWaitingStatus = False
    PlotterClientInitialized = False
    PlotterClientInit = ""
    PlotterClientTERM = ""
    'R E
    RaiseEvent WiFiChange
End Sub
