Private Sub Class_Terminate()
    If PlotterClientInitialized Then DeInitialize
    PlotterClient.Disconnect
End Sub
