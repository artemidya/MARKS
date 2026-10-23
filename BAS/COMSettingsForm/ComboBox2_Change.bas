Private Sub ComboBox2_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub
