Private Sub SendCutCommand()
    Dim cp As Profile
    Set cp = Profiles(TabStrip1.SelectedItem.index)
    Dim ComSettingsStr As String
    ComSettingsStr = OutPuts.ComSettingsStr

    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 52 + LenB(Text4.text) + LenB(KillBirdsAndSpaces(Combo1.text)) + LenB(OutPuts.VidAndPid(KillBirdsAndSpaces(Combo1.text))) + LenB(OutPuts.ExparedID(KillBirdsAndSpaces(Combo1.text))) + LenB(ComSettingsStr)
        .bPut_Long Commands.Cut
        .bPut_Long cp.CommandLanguage
        .bPut_Long cp.ZeroPosition
        .bPut_Long cp.CoordinateSystem
        .bPut_Long cp.LineType
        .bPut_Boolean cp.OneCommandOnPath
        .bPut_Boolean False 'plotter is autoselected
        .bPut_Long Len(Text4.text)
        .bPut_String Text4.text
        .bPut_Long Len(KillBirdsAndSpaces(Combo1.text))
        .bPut_String KillBirdsAndSpaces(Combo1.text)
        .bPut_Long Len(OutPuts.VidAndPid(KillBirdsAndSpaces(Combo1.text)))
        .bPut_String OutPuts.VidAndPid(KillBirdsAndSpaces(Combo1.text))
        .bPut_Long Len(OutPuts.ExparedID(KillBirdsAndSpaces(Combo1.text)))
        .bPut_String OutPuts.ExparedID(KillBirdsAndSpaces(Combo1.text))
        .bPut_Long CLng(cp.WidthScale * 10000)
        .bPut_Long CLng(cp.HeightScale * 10000)
        .bPut_Long Len(ComSettingsStr)
        .bPut_String ComSettingsStr
        
        BlocksSocket1.Send .bytes
    End With
End Sub
