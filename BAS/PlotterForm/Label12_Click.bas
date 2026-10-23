Private Sub Label12_Click()
    Profiles(TabStrip1.SelectedItem.index).SeparateParametrs = Not Profiles(TabStrip1.SelectedItem.index).SeparateParametrs
    RefreshFields
End Sub
