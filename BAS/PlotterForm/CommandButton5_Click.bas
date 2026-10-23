Private Sub CommandButton5_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(2).file, "", 1
10
End Sub
