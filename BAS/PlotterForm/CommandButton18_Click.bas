Private Sub CommandButton18_Click()
    Dim d As Long
    d = ((ValP(TextBox9.text) * 1000 + 24) \ 25) * 25
    d = d - 25
    If d < 0 Then d = 0
    If d > 10000 Then d = 10000
    TextBox9.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox9.SetFocus
    TextBox9.SelStart = 0
    TextBox9.SelLength = 5
End Sub
