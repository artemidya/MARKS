Private Sub CheckBox5_Change()
    Dim np As Boolean: np = Not CheckBox5.value
    TabStrip1.Tabs(4).Visible = np
    TextBox5.Enabled = np
    Label5.Enabled = np
    Profiles(4).Visible = np
End Sub
