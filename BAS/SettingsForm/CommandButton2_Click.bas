Private Sub CommandButton2_Click()
    If MsgBox(localization.ActiveLng.Item("SettingsForm.QuestionDeleteThisCode"), vbYesNo Or vbQuestion, "PlotCalc") = vbYes Then
        UserData.CodeOfActivation = ""
        UserData.Save
        Label4.Caption = ""
        CommandButton2.Enabled = False
        Label6.Enabled = False
        CommandButton1.Enabled = True
        CommandButton1.SetFocus
    End If
End Sub
