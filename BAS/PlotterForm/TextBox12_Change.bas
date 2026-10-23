Private Sub TextBox12_Change()
    Dim txt As String
    txt = Replace(TextBox12.text, vbNewLine, "")
    txt = Trim(txt)
    If Combo1.ListIndex > -1 Then
        If Len(txt) > 0 And OutPuts.Connected(Combo1.ListIndex) = True Then
            CommandButton20.Enabled = True
        Else
            CommandButton20.Enabled = False
        End If
    End If
End Sub
