Private Sub TabStrip1_Change()
    RegData1.ActiveProfile = TabStrip1.value
    StopShowAll
    If TabStrip1.SelectedItem.index = 8 Then
        Frame1.Visible = True
        Frame2.Visible = False
        
        TextBox11.Enabled = False
        Label20.Enabled = False
        TextBox12.Enabled = False
        CommandButton20.Enabled = False
    Else
        Frame1.Visible = False
        Frame2.Visible = True
        If Initialized Then
            RefreshFields
        End If
        
        TextBox11.Enabled = True
        Label20.Enabled = True
        TextBox12.Enabled = True
        TextBox12_Change
        RefreshPltDocument
    End If
End Sub
