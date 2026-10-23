Private Sub OutToWiFi(ByRef bytes() As Byte)
    If (PlotterClientInit <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer _
    Or PlotterClientTERM <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator) Then
        If PlotterClientInitialized Then DeInitialize
        Initialize
    End If
        
    PlotterClient.SendOld bytes()
End Sub
