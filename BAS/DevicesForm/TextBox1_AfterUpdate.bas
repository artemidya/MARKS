Private Sub TextBox1_AfterUpdate()
    On Error GoTo 10
    Dim V As Double
    V = ValP(TextBox1.text) / 100
10
    If V > 5 Then V = 5
    If V >= 0 And V < 0.5 Then V = 0.5
    TextBox1.text = Format(V * 100, "0.00") & "%"
End Sub
