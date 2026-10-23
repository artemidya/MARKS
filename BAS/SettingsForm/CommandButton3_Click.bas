Private Sub CommandButton3_Click()
    If UpdateStatus = GetVersion Then
        CommandButton3.Enabled = False
        Me.Repaint
        Dim version As Long
        Label9.Caption = "connecting.."
        version = UpdateCenter1.version
        If version > -1 Then
            Label9.Caption = left(version, 4) & "." & Mid(version, 5, 2) & "." & right(version, 2)
            Label5.Caption = left(UserData.PlotCalcVersion, 4) & "." & Mid(UserData.PlotCalcVersion, 5, 2) & "." & right(UserData.PlotCalcVersion, 2) & " " & ChrW(40) & ChrW(9685) & ChrW(8255) & ChrW(9685) & ChrW(41)
            UpdateStatus = GetPlotCalcGMS
            RefreshUpdateButton
            If version > UserData.PlotCalcVersion Then
                'Need refresh
                Label5.Font.Bold = True
                Label5.ForeColor = vbRed
                CommandButton3.Enabled = True
                Label9.Caption = left(version, 4) & "." & Mid(version, 5, 2) & "." & right(version, 2)
                Label5.Caption = left(UserData.PlotCalcVersion, 4) & "." & Mid(UserData.PlotCalcVersion, 5, 2) & "." & right(UserData.PlotCalcVersion, 2) & " " & ChrW(40) & ChrW(215) & ChrW(-433) & ChrW(215) & ChrW(41)
            End If
            
        Else
            Label9.Caption = "error"
            CommandButton3.Enabled = True
        End If
    Else
        'Downloading PlotCalc.gms
        CommandButton3.Enabled = False
        Me.Repaint
        Err = False
        UpdateCenter1.UpdatePlotCalc
        If Err = False Then
            MsgBox localization.ActiveLng.Item("SettingsForm.MsgboxSuccess") & vbNewLine & localization.ActiveLng.Item("SettingsForm.MsgboxSuccessEndPart"), vbInformation, "PlotCalc"
            Unload Me
        End If
    End If
End Sub
