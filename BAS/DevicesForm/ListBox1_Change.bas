Private Sub ListBox1_Change()
    If ListBox1.ListIndex < 2 Then
        Command5.Enabled = True
    ElseIf Connected = True Then
        Command5.Enabled = True
    End If
End Sub
