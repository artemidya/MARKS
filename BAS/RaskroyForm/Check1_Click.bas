Private Sub Check1_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    RefreshCommand1
End Sub
