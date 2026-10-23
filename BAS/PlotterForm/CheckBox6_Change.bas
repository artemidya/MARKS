Private Sub CheckBox6_Change()
    Dim np As Boolean: np = Not CheckBox6.value
    TabStrip1.Tabs(5).Visible = np
    TextBox6.Enabled = np
    Label6.Enabled = np
    Profiles(5).Visible = np
End Sub
