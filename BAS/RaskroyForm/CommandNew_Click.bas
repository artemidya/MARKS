Private Sub CommandNew_Click()
    'New profile Button
    
    'proverit Change
    If NeedSave = True Then
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
    NeedSave = False
    
    
    ReDim Preserve UserData.RaskroyProfiles(UserData.CountRaskroyProfiles)
    UserData.CountRaskroyProfiles = UserData.CountRaskroyProfiles + 1

    Dim newName As String
    newName = CreateNewNameForProfile
    UserData.RaskroyProfiles(UserData.CountRaskroyProfiles - 1).name = newName
    Combo1.AddItem newName, UserData.CountRaskroyProfiles - 1

    With UserData.RaskroyProfiles(UserData.CountRaskroyProfiles - 1)
        .DrawMaterial = True
        .DrawWorkObl = True
        .DrawFields = True
        .DrawText = True
        .Group = True
    End With
    

    Combo1.ListIndex = UserData.CountRaskroyProfiles - 1
    Combo1.Enabled = True
    CommandDelete.Enabled = True
    EnebleProfileAllElements
    UserData.Save
    
    Combo1.SetFocus
    Combo1.SelStart = 0
    Combo1.SelLength = 50
    
    NotNest = True
    RefreshCommand1
End Sub
