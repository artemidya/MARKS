Private Sub CheckBox1_Change()
    Dim np As Boolean: np = Not CheckBox1.value
    TabStrip1.Tabs(0).Visible = np
    TextBox1.Enabled = np
    Label1.Enabled = np
    Profiles(0).Visible = np
End Sub
