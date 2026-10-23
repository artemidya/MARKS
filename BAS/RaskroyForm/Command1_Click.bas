Private Sub Command1_Click()
    'Commands.CurveData - отправка данных на сервер
    'Commands.Place - отправка параметров на сервер
    'Commands.Stop - сброс
    'Commands.MoreCombinations  - more combinations
 
    If Processing_Data Then   '[Stop]
        Dim buffer1 As New buffer
        buffer1.InterPlatformNetConverter = True
        buffer1.size = 4
        buffer1.bPut_Long Commands.Stop 'Send [Stop]
        
        BlocksSocket1.Send buffer1.bytes
        Processing_Data = False
        ProtocolStatus = Ready

        Label1.Caption = localization.ActiveLng.Item("RaskroyForm.ActionCancelledByUser")
        ProgressBar1.value = 0

        If ImproveOn Then
            Command1.Caption = localization.ActiveLng.Item("RaskroyForm.ToImprove")
            Command1.BackColor = RGB(147, 244, 186)
            ProgressBar1.GreenColorOn = True
        Else
            Command1.Caption = localization.ActiveLng.Item("RaskroyForm.Command1.Caption")
            Command1.BackColor = VBA.SystemColorConstants.vbButtonFace
            ProgressBar1.GreenColorOn = False
        End If
        EnableAllButtons
        Exit Sub
    End If
    
    If ImproveOn Then   '[MoreCombinations]
        ProtocolStatus = WaitingNestingResult
        Command1.Caption = localization.ActiveLng.Item("RaskroyForm.Stop")
        Processing_Data = True
        Label1.Caption = localization.ActiveLng.Item("RaskroyForm.WaitingServerResponse") '& "    " & localization.ActiveLng.Item("RaskroyForm.Rebuild") & " " & Pp & "   " & localization.ActiveLng.Item("RaskroyForm.Improved") & " " & Improved
        
        DisebleAllButtons
        
        Dim buffer2 As New buffer
        buffer2.InterPlatformNetConverter = True
        buffer2.size = 4
        buffer2.bPut_Long Commands.MoreCombinations 'Send [MoreCombinations]
        BlocksSocket1.Send buffer2.bytes
        
        If SelectedResult < CountResults Or MultiPage1.SelectedItem.name = "Page2" Or SelectedResult > CountResults Then
            SelectedResult = CountResults
            If Plotf = False Then Command4.Enabled = True
            
            Dim wrtemp As Work_Result
            wrtemp = Work_Result(SelectedResult)
            DrawWarkResult wrtemp
        End If
        Exit Sub
    End If
    
    pp = 1
    Command1.Caption = localization.ActiveLng.Item("RaskroyForm.Stop")
    Processing_Data = True
    
    If SelectedResult < CountResults + 1 Or MultiPage1.SelectedItem.name = "Page2" Then
        SelectedResult = CountResults + 1
        Command3.Enabled = False
        Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    End If
    DisebleAllButtons
    If Sending_Data = False Then
        'Отправка данных на сервер
        Label1.Caption = localization.ActiveLng.Item("RaskroyForm.SendingDataServer")
        ProtocolStatus = WaitingDataSentResult
        Command1.Enabled = False
        BlocksSocket1.DeliveryReport = True
        BlocksSocket1.Send CurveData
        BlocksSocket1.DeliveryReport = False
        
    Else
        Label1.Caption = localization.ActiveLng.Item("RaskroyForm.WaitingServerResponse")
        ProtocolStatus = WaitingNestingResult
        StartNesting
    End If
End Sub
