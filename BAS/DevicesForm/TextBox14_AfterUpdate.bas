Private Sub TextBox14_AfterUpdate()
    On Error GoTo 10
    Dim V As Long
    V = Int(ValP(TextBox14.text))
    If V > 100000 Then V = 100000
10
    TextBox14.text = V & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub
