Private Sub CheckBox4_Change()
    Dim np As Boolean: np = Not CheckBox4.value
    TabStrip1.Tabs(3).Visible = np
    TextBox4.Enabled = np
    Label4.Enabled = np
    Profiles(3).Visible = np
End Sub
