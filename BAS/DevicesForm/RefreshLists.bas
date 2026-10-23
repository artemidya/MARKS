Private Sub RefreshLists()
    Dim q As Long
    With ListBox1
        .Clear
        .AddItem "    " & localization.ActiveLng.Item("Profile.Device")
        .AddItem "    " & localization.ActiveLng.Item("Profile.DeviceDMPL")
        If Not Not PlottersList Then
            If DeveloperMode Then
                SortPlottersListByDate
                For q = 0 To UBound(PlottersList)
                    .AddItem DateLngToString(PlottersList(q).Date) & " " & PlottersList(q).Device
                Next
            Else
                For q = 0 To UBound(PlottersList)
                    If DeviceIsPresentInMySystem(PlottersList(q)) Then
                        .AddItem ChrW(10003) & PlottersList(q).Device
                    Else
                        .AddItem "    " & PlottersList(q).Device
                    End If
                Next
            End If
        End If
    End With
    With ListBox2
        .Clear
        If Not Not NeedAddList Then
            For q = 0 To UBound(NeedAddList)
                .AddItem DateLngToString(NeedAddList(q).Date) & " " & NeedAddList(q).Device
            Next
        End If
    End With
End Sub
