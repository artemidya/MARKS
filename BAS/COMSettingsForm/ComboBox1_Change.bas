Private Sub ComboBox1_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub
