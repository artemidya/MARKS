Private Sub TextBox9_AfterUpdate()
    On Error GoTo 10
    Dim d As Long
    d = ((ValP(TextBox9.text) * 1000 + 12.5) \ 25) * 25
    If d < 0 Then d = 0
    If d > 10000 Then d = 10000
10
    TextBox9.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub
