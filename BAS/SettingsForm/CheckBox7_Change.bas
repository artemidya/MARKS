Private Sub CheckBox7_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub
