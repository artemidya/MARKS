Private Sub ForEventsLoop()
    Do While isActive
        For q = 0 To 6
            Sleep 20
            DoEvents
        Next
        If Not PlotterForm.USB_Events1 Is Nothing Then
            PlotterForm.USB_Events1.USB_change_check_loop
        End If
    Loop
End Sub
