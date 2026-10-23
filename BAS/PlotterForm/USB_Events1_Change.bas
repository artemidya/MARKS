Private Sub USB_Events1_Change()
    Dim selItem As String
    Dim selIndex As Long
    
    OutPuts.Refresh
    
    With Combo1
        selItem = .text
        selIndex = .ListIndex
        suppressChange = True
    
        AddOrDeletItemsInCombo OutPuts.count - .ListCount, Combo1
        For q = 0 To Combo1.ListCount - 1
            If OutPuts.Connected(q) = True Then
                If .List(q) <> ChrW(10003) & OutPuts.Names(q) Then
                    .List(q) = ChrW(10003) & OutPuts.Names(q)
                End If
            Else
                If .List(q) <> "    " & OutPuts.Names(q) Then
                    .List(q) = "    " & OutPuts.Names(q)
                End If
            End If
        Next
        suppressChange = False
        If .text <> selItem Then
            .ListIndex = ListIndexByName(KillBirdsAndSpaces(selItem), Combo1, selIndex)
        End If
    End With
    Combo1_Change
    
    If COMSettingsForm.isActive Then
        Unload COMSettingsForm
        OutPuts.Configure KillBirdsAndSpaces(Combo1.text), FindWindow(vbNullString, Me.Caption)
    End If
End Sub
