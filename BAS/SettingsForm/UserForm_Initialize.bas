Private Sub UserForm_Initialize()
    UserData.Load
    
    Combo2.AddItem
    Combo2.ListIndex = 0
    SetTexts

    IsEditingMyFields = True
    
    Combo1.List = Array("RU", "EN", "UK", "ZH", "PL", "DE", "FR", "SK", "SL", "IT", "VI", "KK", "PT", "ES", "TR", "ID", "AR", "HI")
    SelectLangInCombo (UserData.LangCode)
    
    Text4.text = Trim(UserData.PerimetrFormData.MoneyUnit)
    Label4.Caption = UCase(Trim(UserData.CodeOfActivation))

    Dim version As String
    version = left(UserData.PlotCalcVersion, 4) & "." & Mid(UserData.PlotCalcVersion, 5, 2) & "." & right(UserData.PlotCalcVersion, 2)
    Label5.Caption = version
    CheckBox1.value = RegData1.AutoUpdate
    If Len(Trim(UserData.CodeOfActivation)) > 0 Then
        CommandButton2.Enabled = True
        Label6.Enabled = True
    End If
    
    CheckBox2.value = RegData1.DelFill
    CheckBox3.value = RegData1.DelOutline
    CheckBox4.value = RegData1.UngroupAll
    CheckBox5.value = RegData1.OutlineUse
    CheckBox6.value = RegData1.ConnectingCurves
    CheckBox7.value = RegData1.ReduceCurves
    
        
    TextBox1.text = CDbl(RegData1.CropBMPoffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox2.text = CDbl(RegData1.TrimOffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox3.text = CDbl(RegData1.Sensitivity) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm")

    Set UpdateCenter1 = New UpdateCenter
    IsEditingMyFields = False
End Sub
