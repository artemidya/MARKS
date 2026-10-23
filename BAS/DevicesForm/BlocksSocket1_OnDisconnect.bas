Private Sub BlocksSocket1_OnDisconnect(client As BlocksSocket)
    Label40.Caption = localization.ActiveLng.Item("DevicesForm.Label40.Caption_dicnct")
    Connected = False
    Command5.Enabled = False
    CommandButton1.Enabled = False
    Erase PlottersList
    Erase NeedAddList
    RefreshLists
End Sub
