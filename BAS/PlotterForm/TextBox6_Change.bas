Private Sub TextBox6_Change()
    TabStrip1.Tabs(5).Caption = TextBox6.text
    Profiles(5).name = TextBox6.text
End Sub
