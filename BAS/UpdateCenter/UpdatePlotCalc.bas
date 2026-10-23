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
