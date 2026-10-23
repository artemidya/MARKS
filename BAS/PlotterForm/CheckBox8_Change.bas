Private Sub CheckBox8_Change()
    Dim np As Boolean: np = Not CheckBox8.value
    TabStrip1.Tabs(7).Visible = np
    TextBox8.Enabled = np
    Label8.Enabled = np
    Profiles(7).Visible = np
End Sub
