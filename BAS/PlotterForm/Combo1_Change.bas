Private Sub Combo1_Change()
    If TabStrip1.SelectedItem.index > 7 Then Exit Sub
    If suppressChange Then Exit Sub
    Profiles(TabStrip1.SelectedItem.index).OutPut = KillBirdsAndSpaces(Combo1.text)
    RefreshXY
    If Combo1.ListIndex + 1 = Combo1.ListCount Then
        CommandButton22.Enabled = False
        If Not OriginalCurves Is Nothing Then
            If OriginalCurves.count = 0 Then
                Command5.Enabled = False
            Else
                Command5.Enabled = True
            End If
        End If
        Label22.Visible = False

        
        CommandButton21.Enabled = True
        CommandButton11.Enabled = True
        CommandButton13.Enabled = True
        CommandButton12.Enabled = True
        CommandButton14.Enabled = True
        
        Combo1.BackColor = &H80000005
    Else
        If OutPuts.Connected(Combo1.ListIndex) = True Then
            CommandButton22.Enabled = True

            If Not OriginalCurves Is Nothing Then
                If OriginalCurves.count > 0 Then
                    Command5.Enabled = True
                Else
                    Command5.Enabled = False
                End If
            End If

            Label22.Visible = False
            Label23.Visible = False
            
            CommandButton21.Enabled = True
            CommandButton11.Enabled = True
            CommandButton13.Enabled = True
            CommandButton12.Enabled = True
            CommandButton14.Enabled = True
            
            Combo1.BackColor = &H80000005
        Else
            CommandButton22.Enabled = False
            Command5.Enabled = False
            Label22.Visible = True
            
            
            CommandButton21.Enabled = False
            CommandButton11.Enabled = False
            CommandButton13.Enabled = False
            CommandButton12.Enabled = False
            CommandButton14.Enabled = False
            
            Combo1.BackColor = &HFFFF&
        End If
    End If
    Combo1.ControlTipText = KillBirdsAndSpaces(Combo1.text)
    TextBox12_Change
    RefreshStartButton
End Sub
