Private Sub TextBox2_Change()
    TabStrip1.Tabs(1).Caption = TextBox2.text
    Profiles(1).name = TextBox2.text
End Sub
