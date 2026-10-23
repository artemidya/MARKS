Private Sub Label6_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", localization.ActiveLng.Item("SettingsForm.TestCodeLink") & Trim(UserData.CodeOfActivation), "", "", 4
10:
End Sub
