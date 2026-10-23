Private Sub ComboBox4_Change()
    If Initialized Then CommandButton3.Enabled = True
    If ComboBox4.ListIndex = 0 Then
        Label23.Enabled = False
        Label22.Enabled = True
        ComboBox5.Enabled = True
    Else
        Label23.Enabled = True
        Label22.Enabled = False
        ComboBox5.Enabled = False
    End If
End Sub
