Private Sub Text1_Change()
    If Len(Text1.text) = 20 Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
    If IsEditingMyFields Then Exit Sub
    NeedSave = True
End Sub
