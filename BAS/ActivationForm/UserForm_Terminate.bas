Private Sub UserForm_Terminate()
    If NeedSave Then
        UserData.CodeOfActivation = Trim(Text1.text)
        UserData.Save
    End If
End Sub
