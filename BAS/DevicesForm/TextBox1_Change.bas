Private Sub TextBox1_Change()
    If Initialized Then CommandButton3.Enabled = True
    If ValP(TextBox1.text) = 100 And ValP(TextBox2.text) = 100 Then
        CommandButton7.Enabled = False
        Label50.Enabled = False
        Image1.Picture = eu_disabled.Picture
    Else
        CommandButton7.Enabled = True
        Label50.Enabled = True
        Image1.Picture = eu_enabled.Picture
    End If
    
    If ValP(TextBox1.text) = 98.43 And ValP(TextBox2.text) = 98.43 Then
        CommandButton6.Enabled = False
        Label51.Enabled = False
        Image2.Picture = us_disabled.Picture
    Else
        CommandButton6.Enabled = True
        Label51.Enabled = True
        Image2.Picture = us_enabled.Picture
    End If
End Sub
