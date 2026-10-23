Private Sub CommandSave_Click()
    'Save profile Button
    SaveProfile Combo1LastSelectListIndex
    CommandSave.Enabled = False
    NeedSave = False

    'Combo1_Click
End Sub
