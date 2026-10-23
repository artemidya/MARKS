Private Sub DeInitialize()

    On Error GoTo 10
    PlotterForm.AddLog PlotterClientTERM
    Dim bytes() As Byte
    bytes = StrToAscii(PlotterClientTERM)

    PlotterClientInitialized = False
    PlotterClientInit = ""
    PlotterClientTERM = ""

    
    If UBound(bytes) > -1 Then
        PlotterClient.SendOld bytes()
    End If
10
End Sub
