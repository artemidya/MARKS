Private Sub Initialize()
    On Error GoTo 10
    PlotterClientInit = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer
    PlotterClientTERM = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator
    PlotterClientInitialized = True
    
    PlotterForm.AddLog PlotterClientInit
    Dim bytes() As Byte
    bytes = StrToAscii(PlotterClientInit)
    
    If UBound(bytes) > -1 Then
        PlotterClient.SendOld bytes()
    End If
    
10
End Sub
