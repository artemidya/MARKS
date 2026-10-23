Private Sub TextBox16_AfterUpdate()
    On Error GoTo 10
    TextBox16.text = CLng(val(TextBox16.text)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    Exit Sub
10
    TextBox16.text = Profiles(PlotterForm.TabStrip1.SelectedItem.index).LaserDY & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub
