Private Function GetPlottersList(ByRef content As String) As PlotterInfo()
      Dim content2 As String
      content2 = GetTag("PlottersList", content)
      Dim arr() As String
      Dim arr2() As String
      arr = Split(content2, "<br>", , vbTextCompare)
      Dim lGetPlottersList() As PlotterInfo
      
      If UBound(arr) > -1 Then
            ReDim lGetPlottersList(UBound(arr))
      End If
      
      Dim q As Long
      For q = 0 To UBound(arr)
            arr2() = Split(arr(q), ";", , vbTextCompare)
            If UBound(arr2) > 1 Then
                  With lGetPlottersList(q)
                        On Error GoTo 5
                        .Device = arr2(0)
                        .ID = CLng(val(arr2(1)))
                        .Date = CLng(val(arr2(2)))
                        .VidAndPid = arr2(13)
5
                  End With
            End If
      Next
      GetPlottersList = lGetPlottersList
End Function
