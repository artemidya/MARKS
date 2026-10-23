Private Function OutToFile(ByRef bytes() As Byte) As Boolean

    Dim fileExtension As String
    fileExtension = LoadSaveDialog.GetFileExtension(Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT)
    If fileExtension = "" Then fileExtension = "*"
 
    Dim filter As LoadSaveDialog.filter
    With filter
        If UCase(fileExtension) = "PLT" Then
            ReDim .FilterRows(1)
            .FilterRows(0).Mask = "*.PLT"
            .FilterRows(0).Description = "Plotter file (*.plt)"
            .FilterRows(1).Mask = "*.*"
            .FilterRows(1).Description = "All files (*.*)"
        Else
            ReDim .FilterRows(2)
            .FilterRows(0).Mask = "*." & UCase(fileExtension)
            .FilterRows(0).Description = PlotterForm.Text4.text & " (*." & LCase(fileExtension) & ")"
            .FilterRows(1).Mask = "*.PLT"
            .FilterRows(1).Description = "Plotter file (*.plt)"
            .FilterRows(2).Mask = "*.*"
            .FilterRows(2).Description = "All files (*.*)"
        End If
        Dim strFilter As String
        Dim q As Long
        For q = 0 To UBound(.FilterRows)
            strFilter = strFilter & .FilterRows(q).Description & "|" & .FilterRows(q).Mask & "|"
        Next
    End With
    
    Dim result As String
'    result = LoadSaveDialog.SaveFile(Filter, PlotterForm.RegData1.LastPlotterDirectory, Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT)
    On Error GoTo 20
    result = Application.CorelScriptTools.GetFileBox(strFilter, , 1, Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT, , PlotterForm.RegData1.LastPlotterDirectory)
20
    If result <> "" Then
        Profiles(PlotterForm.TabStrip1.SelectedItem.index).FilePLT = LoadSaveDialog.GetFile(result)
        PlotterForm.RegData1.LastPlotterDirectory = LoadSaveDialog.GetPath(result)

        If fFileLen(result) > UBound(bytes) + 1 Then KillFile result
        Dim fileNumber As Long
        fileNumber = FreeFile
        Open result For Binary Access Write As fileNumber
            Put fileNumber, , bytes()
        Close fileNumber
        OutToFile = True
    End If
End Function
