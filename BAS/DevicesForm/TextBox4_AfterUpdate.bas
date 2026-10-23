Private Sub TextBox4_AfterUpdate()
On Error GoTo 10
    If ValP(TextBox4.text) < 10000 Then TextBox4.text = "10000"
    Exit Sub
10
    TextBox4.text = "10000"
End Sub
