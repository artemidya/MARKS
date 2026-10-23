Private Sub CommandButton8_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(5).file, "", 1
10
End Sub
