Private Sub TextBox1_Change()
    TabStrip1.Tabs(0).Caption = TextBox1.text
    Profiles(0).name = TextBox1.text
End Sub
