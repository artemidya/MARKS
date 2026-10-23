Private Sub CommandDelete_Click()
    'Delete profile Button
    If MsgBox(localization.ActiveLng.Item("RaskroyForm.QDeleteProfile") & " (" & Trim(UserData.RaskroyProfiles(Combo1LastSelectListIndex).name) & ")?", vbYesNo + vbQuestion, "PlotCalc") = vbNo Then Exit Sub
    
    RemoveProfile Combo1LastSelectListIndex
    Combo1.RemoveItem Combo1LastSelectListIndex
    NeedSave = False
    If UserData.CountRaskroyProfiles > 0 Then
        If Combo1LastSelectListIndex > UserData.CountRaskroyProfiles - 1 Then
            Combo1.ListIndex = UserData.CountRaskroyProfiles - 1
        Else
            Combo1.ListIndex = Combo1LastSelectListIndex
        End If
    Else
        NotNest = True
        RefreshCommand1
        Combo1.text = " " & localization.ActiveLng.Item("RaskroyForm.empty")
        Combo1.Enabled = False
        CommandDelete.Enabled = False
        CommandSave.Enabled = False
        DisebleProfileAllElements
        If MultiPage1.SelectedItem.name = "Page2" Then
            Cls True
        End If
    End If
End Sub
