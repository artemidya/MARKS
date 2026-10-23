Private Sub FilerRefresh()
    If IsDriveFixed(PlotterForm.RegData1.LastPlotterDirectory) Then
        FileConnectedIfUSB = True
    Else
        FileConnectedIfUSB = FolderExists(PlotterForm.RegData1.LastPlotterDirectory)
    End If
End Sub
