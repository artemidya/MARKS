Private Sub CheckDefaultButton()
    If Combo1.value = 9600 And _
    ComboBox1.value = 8 And _
    ComboBox2.value = "None" And _
    ComboBox3.value = "1" And _
    ComboBox4.value = "None" Then
        Command5.Enabled = False
    Else
        Command5.Enabled = True
    End If
End Sub
