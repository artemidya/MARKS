Private Sub Text14_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub
