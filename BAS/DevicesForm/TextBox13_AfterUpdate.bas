Private Sub TextBox13_AfterUpdate()
    On Error GoTo 10
    Dim V As Long
    V = Int(ValP(TextBox13.text))
    If V > 100000 Then V = 100000
10
    TextBox13.text = V & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub
