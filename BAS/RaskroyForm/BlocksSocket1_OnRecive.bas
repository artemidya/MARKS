Private Sub BlocksSocket1_OnRecive(bytes() As Byte, client As BlocksSocket)
    On Error GoTo 10
    If UBound(bytes) > 2 Then
        Dim buffer1 As New buffer
        Dim command As Commands
        buffer1.InterPlatformNetConverter = True
        buffer1.SetBytes bytes
        command = buffer1.bGet_Long
        
        If command = SetSessionID Then
            SessionID = buffer1.bGet_Long
            Exit Sub
        ElseIf command = ErrorDoubleUseCode Then
            Label1.Caption = "Error Double Use Code " & UserData.CodeOfActivation
            Error = True
            SessionID = 0
            BlocksSocket1.Disconnect
            Exit Sub
        ElseIf command = ErrorSessionNotFound Then
            Label1.Caption = "Session is not found"
            Error = True
            SessionID = 0
            BlocksSocket1.Disconnect
            Exit Sub
        ElseIf command = CloseSession Then
            Label1.Caption = "Session is closed by the PlotCalc server"
            Error = True
            SessionID = 0
            BlocksSocket1.Disconnect
            Exit Sub
        End If
        
        Select Case ProtocolStatus
            Case ProtocolStatusEnm.WaitingForActivationStatus
                If command = TimeBeforeFreeWork Then
                    Dim t As Long
                    t = buffer1.bGet_Long
                    If t = 0 Then
                        Plotf = False
                    Else
                        Plotf = True
                        Time_for_work = t
                        ActivationFormIsShow = True
                        Set ActivationForm.ParentForm = Me
                        ActivationForm.Show 1
                        ActivationFormIsShow = False
                    End If
                    ProtocolStatus = Ready
                    EnableCommand1
                    If Command1.Enabled = True Then Command1.SetFocus
                ElseIf command = ActivationSuccess Then 'Активация успешна
                    MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox1"), vbInformation, "PlotCalc"
                    Plotf = False
                    Unload ActivationForm
                    
                ElseIf command = ActivationCodeNotExists Then 'Не сущетвует такого кода
                    MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox2"), vbCritical, "PlotCalc"
                    
                ElseIf command = ActivationCodeTimeOut Then 'Время действия лицензии истекло
                    MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox3"), vbInformation, "PlotCalc"
                Else
                    Label1.Caption = "Unknown command"

                End If
                Timer1.Enabled = True
            Case ProtocolStatusEnm.WaitingDataSentResult
                If command = ServerReceivedAllDataFlag Then
                    Sending_Data = True
                    Command1.Enabled = True
                    Command1.SetFocus
                    ProtocolStatus = WaitingNestingResult
                    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.WaitingServerResponse")
                    ProgressBar1.value = 0
                    StartNesting
                End If
            Case ProtocolStatusEnm.WaitingNestingResult
                If command = ProgressPercent Then
                    Dim Percent As Long
                    Percent = buffer1.bGet_Long
                    ProgressBar1.value = Percent
                    If ImproveOn Then
                        pp = buffer1.bGet_Long
                        Label10.Caption = localization.ActiveLng.Item("RaskroyForm.Rebuild") & " " & pp & "   " & localization.ActiveLng.Item("RaskroyForm.Improved") & " " & Improved & MmPlus
                    End If
                ElseIf command = result Then
                    If Not ImproveOn Then Label1.Caption = localization.ActiveLng.Item("RaskroyForm.GettingResult")
                    ProgressBar1.value = 0
                    CountResults = CountResults + 1
                    ReDim Preserve Work_Result(1 To CountResults)
                    Work_Result(CountResults) = Get_Work_Result(buffer1) 'Забираем результат
                    SelectedResult = CountResults
                    
                    'Send more
                    If ImproveOn Then
                        Improved = Improved + 1
                        Label10.Caption = localization.ActiveLng.Item("RaskroyForm.Rebuild") & " " & pp & "   " & localization.ActiveLng.Item("RaskroyForm.Improved") & " " & Improved & MmPlus
                    Else
                        Label1.Caption = localization.ActiveLng.Item("RaskroyForm.ResultObtained") & "   (" & Int(Work_Result(SelectedResult).maximumY * 25.4) & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
                        LastMaxY = Work_Result(SelectedResult).maximumY
                        Improved = 0
                        Command1.Caption = localization.ActiveLng.Item("RaskroyForm.ToImprove")
                        Command1.BackColor = RGB(147, 244, 186)
                        ProgressBar1.GreenColorOn = True
                        If Work_Result(SelectedResult).Errors Then Command1.Enabled = False
                        ImproveOn = True
                        EnableAllButtons
                        Processing_Data = False
                        ProtocolStatus = Ready
                    End If
                    'Теперь нарисовать результат
                    Dim wrtemp As Work_Result
                    wrtemp = Work_Result(SelectedResult)
                    DrawWarkResult wrtemp
                    
                ElseIf command = NoMoreCombinations Then
                    Command1.Enabled = False
                    Command1.Caption = localization.ActiveLng.Item("RaskroyForm.ToImprove")
                    Command1.BackColor = VBA.SystemColorConstants.vbButtonFace
                    ProgressBar1.GreenColorOn = False
                    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.impossibleToImproveMore") ' & MmPlus
                    EnableAllButtons
                    Processing_Data = False
                    ProtocolStatus = Ready
                    ProgressBar1.value = 0
                ElseIf command = ErrorContourCreation Then
                    Label1.Caption = "Error Contour Creation"
                ElseIf command = ResultAsBMP Then
                    Label1.Caption = "Result as BMP. Update software"
                    Error = True
                    SessionID = 0
                    BlocksSocket1.Disconnect
                End If
            Case ProtocolStatusEnm.Ready
        End Select
    End If
10
End Sub
