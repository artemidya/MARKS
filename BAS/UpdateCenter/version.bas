Property Get version() As Long
    GetVersionProcessing = True
    UpdateProcessingStatus = WaitingVersion
    
    
    socket1.Connect UserData.MainAddress, UserData.MainAddressPort
    
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.size = 4
    buffer1.bPut_Long Commands.GetCurVersionPlotCalc
    socket1.Send buffer1.bytes

    Dim t As Long
    Do
        Sleep 20
        t = t + 1
        DoEvents
    Loop While UpdateProcessingStatus = WaitingVersion And GetVersionProcessing And t < 200
    version = GettedVersion
    GetVersionProcessing = False
End Property
