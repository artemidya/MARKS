Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If DataHasBeenChanged Then
        Dim res As VbMsgBoxResult
        res = MsgBox(localization.ActiveLng.Item("SettingsForm.QuestionSaveChenges"), vbYesNoCancel Or vbQuestion, "PlotCalc")
        Select Case res
            Case vbYes
                SaveUserData
            Case vbNo
                'Exit without saving
                SelectLangByCode (UserData.LangCode)
                SetAllTexts
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub
