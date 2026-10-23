' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: WiFiSettingsForm.frm
' Procedures: 8
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 12-29
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' Command4_Click (Sub), source lines 30-38
' -----------------------------------------------------------------------------
Private Sub Command4_Click()
    'Save
    On Error GoTo 10
    PlotterForm.RegData1.IP(PlotterForm.TabStrip1.SelectedItem.index + 1) = TextBox1.text & "." & TextBox2.text & "." & TextBox3.text & "." & TextBox4.text
    PlotterForm.RegData1.Port(PlotterForm.TabStrip1.SelectedItem.index + 1) = CInt(TextBox5.text)
    PlotterForm.OutPuts.WiFi_Refresh
10
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' Command5_Click (Sub), source lines 39-41
' -----------------------------------------------------------------------------
Private Sub Command5_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' TextBox1_Change (Sub), source lines 45-48
' -----------------------------------------------------------------------------
Private Sub TextBox1_Change()
    If dntChange Then Exit Sub
    Command4.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox2_Change (Sub), source lines 50-53
' -----------------------------------------------------------------------------
Private Sub TextBox2_Change()
    If dntChange Then Exit Sub
    Command4.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox3_Change (Sub), source lines 55-58
' -----------------------------------------------------------------------------
Private Sub TextBox3_Change()
    If dntChange Then Exit Sub
    Command4.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox4_Change (Sub), source lines 60-63
' -----------------------------------------------------------------------------
Private Sub TextBox4_Change()
    If dntChange Then Exit Sub
    Command4.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox5_Change (Sub), source lines 65-68
' -----------------------------------------------------------------------------
Private Sub TextBox5_Change()
    If dntChange Then Exit Sub
    Command4.Enabled = True
End Sub
