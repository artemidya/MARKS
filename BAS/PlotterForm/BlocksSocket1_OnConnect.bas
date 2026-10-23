Private Sub BlocksSocket1_OnConnect(client As BlocksSocket)
    'Send Head
    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 64
        .bPut_Long Commands.HEAD_CutWindow_20250528 '4
        .bPut_Long UserData.PlotCalcVersion         '4
        .bPut_Long RegData1.FirstCutDateLng         '4
'        .bPut_Long 20250510
        .bPut_Long OriginalCurves.count             '4


        If Application.VersionMajor = 24 And Application.VersionMinor >= 3 Then
            'Corel 2023
            .bPut_Long Application.VersionMajor + 1999 '4
        Else
            If Application.VersionMajor >= 25 Then
                'Corel 2024 +
                .bPut_Long Application.VersionMajor + 1999 '4
            Else
                .bPut_Long Application.VersionMajor + 1998 '4
            End If
        End If
        

        Dim lcode As String * 2
        lcode = UserData.LangCode
        .bPut_String lcode '4
        Dim code As String * 20
        code = Trim(UserData.CodeOfActivation)
        If Len(code) <> 20 Then code = "--------------------"
        .bPut_String code '40

        BlocksSocket1.Send .bytes
        
    End With
End Sub
