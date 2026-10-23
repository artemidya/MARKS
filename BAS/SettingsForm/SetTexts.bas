Private Sub SetTexts()
    Me.Caption = localization.ActiveLng.Item("SettingsForm.Caption")
    Label6.Caption = localization.ActiveLng.Item("SettingsForm.Label6.Caption")
    Label6.left = 91 - Label6.Width \ 2
    Label6.ControlTipText = localization.ActiveLng.Item("SettingsForm.TestCodeLink") & UserData.CodeOfActivation
    
    Frame1.Caption = localization.ActiveLng.Item("SettingsForm.Frame1.Caption")
    Label1.Caption = localization.ActiveLng.Item("SettingsForm.Label1.Caption")
    Label2.Caption = localization.ActiveLng.Item("SettingsForm.Label2.Caption")
    Label3.Caption = localization.ActiveLng.Item("SettingsForm.Label3.Caption")
    Frame3.Caption = localization.ActiveLng.Item("SettingsForm.Frame3.Caption")
    Frame2.Caption = localization.ActiveLng.Item("SettingsForm.Frame2.Caption")
    CommandButton2.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton2.Caption")
    CommandButton1.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton1.Caption")
    Command5.Caption = localization.ActiveLng.Item("SettingsForm.Command5.Caption")
    Combo2.List(0) = localization.ActiveLng.Item("SettingsForm.Combo2.Item1")
    
    Label7.Caption = localization.ActiveLng.Item("SettingsForm.Label7")
    Label8.Caption = localization.ActiveLng.Item("SettingsForm.Label8")
    CheckBox1.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox1.Caption")
    
    Frame4.Caption = localization.ActiveLng.Item("SettingsForm.AdditionalParameters")
    Label10.Caption = localization.ActiveLng.Item("SettingsForm.OffsetCropping")
    Label11.Caption = localization.ActiveLng.Item("SettingsForm.OffsetTrimming")
    
    Label12.Caption = localization.ActiveLng.Item("SettingsForm.Label12")
    CheckBox2.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox2")
    CheckBox3.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox3")
    CheckBox4.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox4")

    Label13.Caption = localization.ActiveLng.Item("SettingsForm.Label13")
    CheckBox5.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox5")
    
    CheckBox6.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox6")
    CheckBox7.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox7")
    
    

    Dim t As Boolean
    t = IsEditingMyFields
    IsEditingMyFields = True
    TextBox1.text = ValP(TextBox1.text) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox2.text = ValP(TextBox2.text) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox3.text = ValP(TextBox3.text) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    IsEditingMyFields = t
    
    RefreshUpdateButton
End Sub
