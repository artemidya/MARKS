Private Sub SaveUserData()
    'Save changes....
    RegData1.AutoUpdate = CheckBox1.value
    RegData1.CropBMPoffset = LngP(ValP(TextBox1.text) * 1000)
    RegData1.TrimOffset = LngP(ValP(TextBox2.text) * 1000)
    RegData1.Sensitivity = LngP(ValP(TextBox3.text) * 1000)
    
    RegData1.DelFill = CheckBox2.value
    RegData1.DelOutline = CheckBox3.value
    RegData1.UngroupAll = CheckBox4.value
    
    RegData1.OutlineUse = CheckBox5.value
    RegData1.ConnectingCurves = CheckBox6.value
    RegData1.ReduceCurves = CheckBox7.value
    
    UserData.PerimetrFormData.MoneyUnit = Text4.text
    UserData.LangCode = Combo1.value
    ToolBar.SetToolTipText
    UserData.Save
    SetAllTexts
End Sub
