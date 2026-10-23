Private Sub CommandButton2_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & TextBox12.text, "", 1
10
End Sub
