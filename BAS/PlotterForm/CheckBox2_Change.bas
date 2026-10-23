Private Sub CheckBox2_Change()
    Dim np As Boolean: np = Not CheckBox2.value
    TabStrip1.Tabs(1).Visible = np
    TextBox2.Enabled = np
    Label2.Enabled = np
    Profiles(1).Visible = np
End Sub
