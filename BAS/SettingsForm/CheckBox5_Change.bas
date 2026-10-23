Private Sub CheckBox5_Change()
    If CheckBox5.value Then
        Label13.Enabled = False
        TextBox3.Enabled = False
    Else
        Label13.Enabled = True
        TextBox3.Enabled = True
    End If

    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub
