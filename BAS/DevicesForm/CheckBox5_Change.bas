Private Sub CheckBox5_Change()
    If Initialized Then CommandButton3.Enabled = True
    If CheckBox5.value Then
        Label44.Enabled = True
        Label45.Enabled = True
        TextBox15.Enabled = True
        TextBox16.Enabled = True
    Else
        Label44.Enabled = False
        Label45.Enabled = False
        TextBox15.Enabled = False
        TextBox16.Enabled = False
    End If
End Sub
