Private Sub RefreshUpdateButton()
    If UpdateStatus = GetVersion Then
        CommandButton3.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton3.CaptionCheck")
    Else
        CommandButton3.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton3.CaptionRefresh")
    End If
'
End Sub
