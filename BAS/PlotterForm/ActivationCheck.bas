Private Sub ActivationCheck()
    If Not PlotCalcIsActive Then
        Time_for_work = Time_for_work_recieved - DateDiff("s", Time_stamp_recived_time, Now)
        If Time_for_work > 0 Then
            ActivationFormIsShow = True
            Set ActivationForm.ParentForm = Me
            If Not BlocksSocket1.Connected Then
                ActivationForm.Command1.Enabled = False
                ActivationForm.Text1.Enabled = False
            End If
            ActivationForm.Show 1
            ActivationFormIsShow = False
        Else
            PlotCalcIsActive = True
        End If
    Else
        If BlocksSocket1.Connected And OriginalCurves.count > 20 And RegData1.FirstCutDateLng = 0 Then
            RegData1.FirstCutDateLng = DateToLong(Now)
        End If
    End If

End Sub
