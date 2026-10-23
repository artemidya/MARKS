Private Function DownloadPlotCalcGMS() As Boolean
    GetPlotCalcGMSProcessing = True
    RaiseEvent DownloadPercent(0)
    UpdateProcessingStatus = DownloadingPlotCalcGMS
    
    socket1.Connect UserData.MainAddress, UserData.MainAddressPort
    
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.size = 4
    buffer1.bPut_Long Commands.GetPlotCalcGMS
    socket1.Send buffer1.bytes
    
    tt = 0
    
    Do
        Sleep 20
        tt = tt + 1
        DoEvents
    Loop While UpdateProcessingStatus = DownloadingPlotCalcGMS And GetPlotCalcGMSProcessing And tt < 200
    
    If UpdateProcessingStatus = PlotCalcGMSReceived Then DownloadPlotCalcGMS = True
    
    GetPlotCalcGMSProcessing = False
End Function
