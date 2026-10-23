Private Sub CheckBox7_Change()
    Dim np As Boolean: np = Not CheckBox7.value
    TabStrip1.Tabs(6).Visible = np
    TextBox7.Enabled = np
    Label7.Enabled = np
    Profiles(6).Visible = np
End Sub
