Private Sub Combo1_Click()
    If Combo1DontClick = True Then Exit Sub

    'proverit Change
    If NeedSave = True Then
        If Combo1.ListIndex = Combo1LastSelectListIndex Then
            Combo1DontClick = True
            Combo1.text = Combo1_Text
            Combo1DontClick = False
            Exit Sub
        End If
    
        Dim selectedAnswer As Variant
        selectedAnswer = MsgBox(localization.ActiveLng.Item("RaskroyForm.QSaveChanges") & " (" & Combo1_Text & ")?", vbYesNoCancel + vbQuestion, "PlotCalc")
        If selectedAnswer = vbYes Then
            SaveProfile Combo1LastSelectListIndex
        ElseIf selectedAnswer = vbCancel Then
            Combo1DontClick = True
            Combo1.text = Combo1_Text
            Combo1DontClick = False
            Exit Sub
        End If
    End If
    
    
    Combo1_Text = Combo1.value
    Combo1LastSelectListIndex = Combo1.ListIndex
    CommandSave.Enabled = False
    NeedSave = False
    
    TextFieldsIsEditing = True
    With UserData.RaskroyProfiles(Combo1.ListIndex)
        Text4.text = .MaterialWith / 1000
        Text5.text = .MaterialHeight / 1000
        Check2.value = .Vnutri_Figuri
        Text8.text = .PlotterSheetWith / 1000
        Text7.text = .PlotterSheetHeight / 1000
        Text12.text = .UpField / 1000
        Text13.text = .LeftField / 1000
        Text14.text = .RightField / 1000
        Text15.text = .DownField / 1000
        Check1.value = .Vlogen
        If .GradusCount = 1 Then
            Combo2.ListIndex = 0
        ElseIf .GradusCount = 2 Then
            Combo2.ListIndex = 1
        ElseIf .GradusCount = 4 Then
            Combo2.ListIndex = 2
        ElseIf .GradusCount = 8 Then
            Combo2.ListIndex = 3
        Else
            Combo2.ListIndex = 4
        End If
        Text18.text = .Otstup / 1000
        
        Check4.value = .DrawMaterial
        Check3.value = .DrawWorkObl
        Check5.value = .DrawFields
        Check6.value = .DrawText
        Check7.value = .PlotCalcLayer
        Check8.value = .Group
        Check9.value = .LockObj
    End With
    TextFieldsIsEditing = False
    
    If MultiPage1.SelectedItem.name = "Page2" Then
        DrawRazm ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    Else
        Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    End If
    RefreshCommand1
End Sub
