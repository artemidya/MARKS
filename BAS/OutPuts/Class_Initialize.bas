Private Sub Class_Initialize()
    'Init BlocksSokets
    BlocksSocketShared.HookForm GetForegroundWindow
    Set PlotterClient = New BlocksSocket
    PlotterClient.OnReciveOld_TurnON = True
    PlotterClientLastTabStrip1IndexValue = -1

    FilerRefresh

    Set printers = New printers
    Set Ports = New Ports
    Set Interfaces = New Interfaces
End Sub
