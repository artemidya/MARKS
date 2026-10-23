Private Sub TextBox3_Change()
    TabStrip1.Tabs(2).Caption = TextBox3.text
    Profiles(2).name = TextBox3.text
End Sub
