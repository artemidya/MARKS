Public Sub WiFi_change_check_loop()
    'Check wifi
    If PlotterClientLastTabStrip1IndexValue <> PlotterForm.TabStrip1.SelectedItem.index Then
        PlotterClient.Disconnect
        PlotterClientlastOperationTime = Timer
        PlotterClientLastTabStrip1IndexValue = PlotterForm.TabStrip1.SelectedItem.index
        Exit Sub
    End If

    If PlotterClientWaitingStatus And (Timer - PlotterClientlastOperationTime) > 6 Then
        PlotterClient.Disconnect
        PlotterClientlastOperationTime = Timer
        Exit Sub
    End If

    '
    If PlotterClientWaitingStatus = False And PlotterClientConnected = False Then
        PlotterClientWaitingStatus = True
        PlotterClientlastOperationTime = Timer
        PlotterClient.Connect PlotterForm.RegData1.IP(PlotterForm.TabStrip1.SelectedItem.index + 1), PlotterForm.RegData1.Port(PlotterForm.TabStrip1.SelectedItem.index + 1)
    End If
End Sub
