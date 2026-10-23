Private Sub UserForm_Activate()
    Label1.Caption = localization.ActiveLng.Item("ActivationForm.Label1.Caption1") & " " & HoursAndMinets(ParentForm.Time_for_work \ 60 + 1) & localization.ActiveLng.Item("ActivationForm.Label1.Caption2") & vbNewLine & vbNewLine _
                    & localization.ActiveLng.Item("ActivationForm.Label1.Caption3")
End Sub
