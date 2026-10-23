Private Sub Combo1_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
    SelectLangByCode (Combo1.value)
    SetTexts
    SetAllTexts
End Sub
