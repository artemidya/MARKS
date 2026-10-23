Private Sub Label2_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", localization.ActiveLng.Item("ActivationForm.BuyURL"), "", "", 4
10:
End Sub
