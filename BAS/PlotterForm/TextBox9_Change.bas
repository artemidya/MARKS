Private Sub TextBox9_Change()
    On Error GoTo 10
    Profiles(TabStrip1.SelectedItem.index).ClosingCurve = ValP(TextBox9.text) * 1000
10
End Sub
