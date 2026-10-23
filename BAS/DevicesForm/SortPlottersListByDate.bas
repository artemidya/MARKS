Private Sub SortPlottersListByDate()
    If Not Not PlottersList Then
        Dim i As Long
        Dim fSwitched As Boolean
        Dim plotterTemp As PlotterInfo
        Do
            fSwitched = False
            For i = 0 To UBound(PlottersList) - 1
                If PlottersList(i).Date < PlottersList(i + 1).Date Then
                    fSwitched = True
                    plotterTemp = PlottersList(i)
                    PlottersList(i) = PlottersList(i + 1)
                    PlottersList(i + 1) = plotterTemp
                End If
            Next
        Loop While fSwitched
    End If
End Sub
