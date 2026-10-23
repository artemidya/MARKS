Private Sub CommandButton16_Click()
    Dim d As Long
    d = ((ValP(TextBox10.text) * 1000) \ 25) * 25
    d = d + 25
    If d < 0 Then d = 0
    If d > 5550 Then d = 5550
    TextBox10.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox10.SetFocus
    TextBox10.SelStart = 2
    TextBox10.SelLength = 3
End Sub
