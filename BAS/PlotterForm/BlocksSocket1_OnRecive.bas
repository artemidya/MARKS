Private Sub BlocksSocket1_OnRecive(bytes() As Byte, client As BlocksSocket)
    On Error GoTo 10
    If UBound(bytes) > 2 Then
        Dim buffer1 As New buffer
        Dim command As Commands
        buffer1.InterPlatformNetConverter = True
        buffer1.SetBytes bytes
        command = buffer1.bGet_Long
        Select Case command
            Case Commands.TimeBeforeFreeWork
                Time_for_work_recieved = buffer1.bGet_Long
                Time_stamp_recived_time = Now
                If Time_for_work_recieved > 0 Then
                    PlotCalcIsActive = False
                Else
                    PlotCalcIsActive = True
                End If
                
            Case Commands.ActivationSuccess
                MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox1"), vbInformation, "PlotCalc"
                PlotCalcIsActive = True
                Unload ActivationForm
                
            Case Commands.ActivationCodeNotExists
                MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox2"), vbCritical, "PlotCalc"
                
            Case Commands.ActivationCodeTimeOut
                MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox3"), vbInformation, "PlotCalc"
                
            Case Else
                'Unknown command
                
        End Select
    End If
    Exit Sub
10
End Sub
