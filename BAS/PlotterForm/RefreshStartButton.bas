Private Sub RefreshStartButton()
    If Combo1.ListIndex + 1 = Combo1.ListCount Then
        Command5.Caption = localization.ActiveLng.Item("PlotterForm.Command5.Caption_SaveAs")
        CheckBox9.Enabled = False
    Else
        If CheckBox9.value Then
            Command5.Caption = localization.ActiveLng.Item("PlotterForm.Command5.Caption_Pre-run")
        Else
            Command5.Caption = localization.ActiveLng.Item("PlotterForm.Command5.Caption_Run")
        End If

        CheckBox9.Enabled = True

    End If
End Sub
