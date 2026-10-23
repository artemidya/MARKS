Private Sub BlocksSocket1_OnConnect(client As BlocksSocket)
    Label11.Visible = True
    Connected = True

    'Send Head
    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 60
        .bPut_Long Commands.HEAD_Raskroy_20220421
        
        .bPut_Long UserData.PlotCalcVersion
        .bPut_Long SessionID
        If Application.VersionMajor = 24 And Application.VersionMinor >= 3 Then
            'Corel 2023
            .bPut_Long Application.VersionMajor + 1999
        Else
            If Application.VersionMajor >= 25 Then
                'Corel 2024 +
                .bPut_Long Application.VersionMajor + 1999
            Else
                .bPut_Long Application.VersionMajor + 1998
            End If
        End If
'        application.corelv
        Dim lcode As String * 2
        lcode = UserData.LangCode
        .bPut_String lcode
        Dim code As String * 20
        code = Trim(UserData.CodeOfActivation)
        If Len(code) <> 20 Then code = "--------------------"
        .bPut_String code
        
        BlocksSocket1.Send .bytes
    End With
    If SessionID <> 0 Then
        ProtocolStatus = RepairProtocolStatus
        Label1.Caption = LastCaption
    Else
        ProtocolStatus = WaitingForActivationStatus
        Label1.Caption = ""
    End If
    
    If ActivationFormIsShow Then
        
        ActivationForm.Text1.Enabled = True
        If Len(ActivationForm.Text1.text) = 20 Then
            ActivationForm.Command1.Enabled = True
        End If
    End If
End Sub
