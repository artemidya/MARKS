Private Sub UserForm_Initialize()
    dntChange = True
    Command4.Caption = localization.ActiveLng.Item("RaskroyForm.Command4.Caption")
    Command5.Caption = localization.ActiveLng.Item("RaskroyForm.Command5.Caption")
    Label5.Caption = localization.ActiveLng.Item("WiFiSettingsForm.Label5.Caption")
    Label6.Caption = localization.ActiveLng.Item("WiFiSettingsForm.Label6.Caption")
    
    Dim ipParts() As String
    ipParts = Split(PlotterForm.RegData1.IP(PlotterForm.TabStrip1.SelectedItem.index + 1), ".")
    If UBound(ipParts) = 3 Then
        TextBox1.text = ipParts(0)
        TextBox2.text = ipParts(1)
        TextBox3.text = ipParts(2)
        TextBox4.text = ipParts(3)
    End If
    TextBox5.text = PlotterForm.RegData1.Port(PlotterForm.TabStrip1.SelectedItem.index + 1)
    dntChange = False
End Sub
