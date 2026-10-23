Private Sub BlocksSocket1_OnConnect(client As BlocksSocket)
    Label40.Caption = localization.ActiveLng.Item("DevicesForm.Label40.Caption_con")
    Connected = True
    BlocksSocket1.Send GetList
    CommandButton1.Enabled = True
End Sub
