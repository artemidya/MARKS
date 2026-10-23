Private Sub KeyReadLoop()
    On Error Resume Next
    KeyLoopProcessing = True
    Dim l As Double:  l = 5
    Dim pKey As PressedKeys
    Static lastTimer As Single
    Static lastTimer2 As Single
    
    Do While KeyLoopProcessing
        pKey = PressedKey
        If ListenKeyboard Then
            If Timer - lastTimer > 0.1 Then
                If CommandButton23.Picture = 0 Then
                    CommandButton23.Picture = Image3.Picture
                Else
                    CommandButton23.Picture = Nothing
                End If
                lastTimer = Timer
            End If
            Select Case pKey
                Case KeyLeft
                    OutPuts.MoveLeft l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyRight
                    OutPuts.MoveRight l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyUp
                    OutPuts.MoveUP l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyDown
                    OutPuts.MoveDown l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyLeftUp
                    OutPuts.MoveLeftUp l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyUpRight
                    OutPuts.MoveUpRight
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyRightDown
                    OutPuts.MoveRightDown
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyDownLeft
                    OutPuts.MoveDownLeft
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyEnter
                    
                Case Key0
                    OutPuts.Move00
                    RefreshXY
                Case Else
                    l = 5
            End Select
        End If

        If Not USB_Events1 Is Nothing Then
            USB_Events1.USB_change_check_loop
            'Chek wifi connection
            OutPuts.WiFi_change_check_loop
        End If
        
        If Label22.Visible Then
            If Timer - lastTimer2 > 0.1 Then
                Label23.Visible = Not Label23.Visible
                lastTimer2 = Timer
            End If
        Else
            Label23.Visible = False
        End If


        For q = 0 To 6
            Sleep 20
            DoEvents
        Next
    Loop
End Sub
