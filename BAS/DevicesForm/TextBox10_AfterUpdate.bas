Private Sub TextBox10_AfterUpdate()
    TextBox10.text = Replace(TextBox10.text, ",", ";")
End Sub
