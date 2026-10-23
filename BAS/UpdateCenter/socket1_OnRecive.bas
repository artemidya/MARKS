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
