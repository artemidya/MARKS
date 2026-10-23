Private Sub BlocksSocket1_OnConnectionProcessIsStarted()
    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.connectTo") & " " & BlocksSocket1.Address & ".."
End Sub
