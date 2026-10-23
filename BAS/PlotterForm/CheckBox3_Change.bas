Private Sub CheckBox3_Change()
    Dim np As Boolean: np = Not CheckBox3.value
    TabStrip1.Tabs(2).Visible = np
    TextBox3.Enabled = np
    Label3.Enabled = np
    Profiles(2).Visible = np
End Sub
