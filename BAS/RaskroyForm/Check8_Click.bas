Private Sub Check8_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
End Sub
