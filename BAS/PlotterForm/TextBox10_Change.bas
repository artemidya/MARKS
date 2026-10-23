Private Sub TextBox10_Change()
    On Error GoTo 10
    Profiles(TabStrip1.SelectedItem.index).KnifeOffset = ValP(TextBox10.text) * 1000
10
End Sub
