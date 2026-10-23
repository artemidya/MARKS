Private Sub Combo1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Combo1_Text <> Combo1.value Then
        CommandSave.Enabled = True
        NeedSave = True
        Combo1_Text = Combo1.value
    End If
End Sub
