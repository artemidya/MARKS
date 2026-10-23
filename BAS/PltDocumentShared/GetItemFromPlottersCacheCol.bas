Private Function GetItemFromPlottersCacheCol(Key As String) As Plotter2DModelCache
    On Error GoTo 10
    Set GetItemFromPlottersCacheCol = PlottersCacheCol(Key)
    Exit Function
10
    Set GetItemFromPlottersCacheCol = Nothing
End Function
