Private Sub CommandButton6_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(3).file, "", 1
10
End Sub
