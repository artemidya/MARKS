Private Sub TextBox10_AfterUpdate()
    On Error GoTo 10
    Dim d As Long
    d = ((ValP(TextBox10.text) * 1000 + 12.5) \ 25) * 25
    If d < 0 Then d = 0
    If d > 5550 Then d = 5550
10
    TextBox10.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub
