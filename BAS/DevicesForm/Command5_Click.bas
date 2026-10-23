Private Sub Command5_Click()
    If ListBox1.ListIndex = 0 Then
        If MsgBox(localization.ActiveLng.Item("DevicesForm.MsgBox.Repl"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then
            SetHPGLstandart
            Frame4.ScrollTop = 0
        End If
    ElseIf ListBox1.ListIndex = 1 Then
        If MsgBox(localization.ActiveLng.Item("DevicesForm.MsgBox.Repl"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then
            SetDMPLstandart
            Frame4.ScrollTop = 0
        End If
    ElseIf ListBox1.ListIndex > 1 Then
        'Internet
        LoadPlotterProcess = True
        TempLabel40 = Label40.Caption
        Dim ID As Long
        ID = PlottersList(ListBox1.ListIndex - 2).ID
        Dim buffer1 As New buffer
        With buffer1
            .InterPlatformNetConverter = True
            .size = 8
            .bPut_Long Commands.GetPlotter
            .bPut_Long ID
            BlocksSocket1.Send .bytes
        End With
    End If
End Sub
