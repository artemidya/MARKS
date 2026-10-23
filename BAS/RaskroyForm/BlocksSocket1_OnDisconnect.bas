Private Sub BlocksSocket1_OnDisconnect(client As BlocksSocket)
    Connected = False
    If ActivationFormIsShow Then
        ActivationForm.Command1.Enabled = False
        ActivationForm.Text1.Enabled = False
    End If
    
    If SessionID <> 0 Then
        'Repair Session
        RepairProtocolStatus = ProtocolStatus
        LastCaption = Label1.Caption
        Set BlocksSocket1 = Nothing
        Set BlocksSocket1 = New BlocksSocket
        BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
    Else
        'End
        Label11.Visible = False
        ProgressBar1.value = 0
        Command1.Enabled = False
        Timer1.Enabled = False
        ProtocolStatus = NotConnected
        If Not Error Then
'            Label1.Font.Bold = True
'            Label1.ForeColor = RGB(240, 68, 77)
            Label1.Caption = localization.ActiveLng.Item("RaskroyForm.NoConnectionServer") & " (" & BlocksSocket1.Address & ")"
        End If
    End If
End Sub
