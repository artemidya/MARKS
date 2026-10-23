Private Sub UserForm_Initialize()
    IsEditingMyFields = True

    Caption = localization.ActiveLng.Item("ActivationForm.Caption")
    Command1.Caption = localization.ActiveLng.Item("ActivationForm.Command1.Caption")
    Label2.Caption = localization.ActiveLng.Item("ActivationForm.Label2.Caption")
    Label2.left = 158 - Label2.Width \ 2
    Label2.ControlTipText = localization.ActiveLng.Item("ActivationForm.BuyURL")
    
    Text1.text = UCase(Trim(UserData.CodeOfActivation))
    IsEditingMyFields = False
End Sub
