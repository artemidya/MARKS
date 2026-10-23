Private Sub ListBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.value = 107 And LastKeyCode = 107 Then
        DeveloperMode = True
        TextBox17.Visible = True
        Me.Caption = "Developer Mode"
        Text4.Locked = False
        Text4.BackStyle = fmBackStyleOpaque
        CommandButton3.Caption = "Send"
        Frame13.Visible = True
        RefreshLists
    End If
    LastKeyCode = KeyCode.value
End Sub
