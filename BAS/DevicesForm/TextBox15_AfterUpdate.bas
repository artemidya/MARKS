Private Sub TextBox15_AfterUpdate()
    On Error GoTo 10
    TextBox15.text = CLng(val(TextBox15.text)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    Exit Sub
10
    TextBox15.text = Profiles(PlotterForm.TabStrip1.SelectedItem.index).LaserDX & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub
