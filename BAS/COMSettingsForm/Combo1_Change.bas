Private Sub Combo1_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub
