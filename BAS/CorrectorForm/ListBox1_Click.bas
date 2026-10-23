Private Sub ListBox1_Click()
    Label2.Caption = ActiveLng.Key(ListBox1.ListIndex)
    TextBox1.text = ActiveLng.value(ListBox1.ListIndex)
    TextBox1.SetFocus
    TextBox1.SelStart = 0
    TextBox1.SelLength = Len(TextBox1.text)
    If CheckBox2.value Then SendKeys "^q"
End Sub
