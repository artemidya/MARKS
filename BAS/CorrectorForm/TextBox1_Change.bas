Private Sub TextBox1_Change()
    ListBox1.List(ListBox1.ListIndex, 1) = TextBox1.text
    localization.ActiveLng.Item(ActiveLng.Key(ListBox1.ListIndex)) = TextBox1.text
    TextBox2.text = ""
End Sub
