Private Sub CheckBox10_Change()
    Profiles(TabStrip1.SelectedItem.index).DoublePass = CheckBox10.value
    CommandButton18.Enabled = Not CheckBox10.value
    CommandButton19.Enabled = Not CheckBox10.value
    TextBox9.Enabled = Not CheckBox10.value
End Sub
