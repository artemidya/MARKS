Private Sub UserForm_Initialize()
    UserData.Load
    If UserData.NumberOfCopies < 1 Then
        Text1.text = 10
    Else
        Text1.text = UserData.NumberOfCopies
    End If
    Text1.SelStart = 0
    Text1.SelLength = 10
    
    If ActiveSelection.shapes.count < 1 Then
        MsgBox localization.ActiveLng.Item("CirculationForm.MsgBox"), vbOKOnly Or vbInformation, "PlotCalc"
        Unload Me
    End If
    
    Command4.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton1.Caption")
    Command5.Caption = localization.ActiveLng.Item("SettingsForm.Command5.Caption")
    
    Label1.Caption = localization.ActiveLng.Item("CirculationForm.Label1.Caption")
    
End Sub
