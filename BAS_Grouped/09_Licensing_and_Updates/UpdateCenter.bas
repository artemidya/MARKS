' =============================================================================
' Group: 09_Licensing_and_Updates
' Source module: UpdateCenter.cls
' Procedures: 7
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 35-39
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    BlocksSocketShared.HookForm GetActiveWindow
    Set socket1 = New BlocksSocket
    GettedVersion = -1
End Sub

' -----------------------------------------------------------------------------
' version (PropertyGet), source lines 40-61
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' UpdatePlotCalc (Sub), source lines 62-85
' -----------------------------------------------------------------------------
Public Sub UpdatePlotCalc()
    On Error GoTo 10
    If DownloadPlotCalcGMS Then
        Dim file As String
        file = VGCore.GMSManager.Projects("PlotCalc").FullFileName
        
'        Replace PlotCalc.gms
        Kill file
        Dim fileNumber As Long
        fileNumber = FreeFile
        Open file For Binary Access Write As fileNumber
            Put fileNumber, , GettedBytes
        Close fileNumber

'        Set socket1 = Nothing
'        VGCore.GMSManager.Projects.Load file
        
    Else
        RaiseEvent Error
    End If
    Exit Sub
10
    RaiseEvent Error
End Sub

' -----------------------------------------------------------------------------
' DownloadPlotCalcGMS (Function), source lines 86-110
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' socket1_OnRecive (Sub), source lines 111-137
' -----------------------------------------------------------------------------
Private Sub socket1_OnRecive(bytes() As Byte, client As BlocksSocket)
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.SetBytes bytes
    Dim command As Commands
    command = buffer1.bGet_Long
    Select Case command
        Case Commands.SetCurVersionPlotCalc
            GettedVersion = buffer1.bGet_Long
            UpdateProcessingStatus = VersionRecieved
            socket1.Disconnect
            
        Case Commands.SetCRC32Signature
            GettedCRC32 = buffer1.bGet_Long
        
        Case Commands.SetPlotCalcGMS
            GettedBytes = buffer1.bGet_Bytes(UBound(bytes) + 1 - 4)
            'Test CRC
            If buffer1.CRC32 = GettedCRC32 Then
                UpdateProcessingStatus = PlotCalcGMSReceived
            Else
                UpdateProcessingStatus = PlotCalcGMSReceivedCRCError
                MsgBox "Incorrect Signature of PlotCalc.gms", vbCritical
            End If
            socket1.Disconnect
    End Select
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 138-142
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    If Not socket1 Is Nothing Then
        socket1.Disconnect
    End If
End Sub

' -----------------------------------------------------------------------------
' socket1_OnStatisticChange (Sub), source lines 144-149
' -----------------------------------------------------------------------------
Private Sub socket1_OnStatisticChange(client As BlocksSocket)
    If UpdateProcessingStatus = DownloadingPlotCalcGMS Then
        tt = 0
        RaiseEvent DownloadPercent(client.Statistic_PercentReceiving)
    End If
End Sub
