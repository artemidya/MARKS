Private Sub Text18_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    RefreshCommand1
End Sub
