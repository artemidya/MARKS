Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If NeedSave Then
        Dim selectedAnswer As Variant
        selectedAnswer = MsgBox(localization.ActiveLng.Item("RaskroyForm.QSaveChanges") & " (" & Combo1_Text & ")?", vbYesNoCancel Or vbQuestion, "PlotCalc")
        If selectedAnswer = vbYes Then
            SaveProfile Combo1LastSelectListIndex
        ElseIf selectedAnswer = vbCancel Then
            Cancel = 1
            OkPress = False
            Exit Sub
        End If
    End If
    Timer1.StopTimer
    '[CloseSession]
    SessionID = 0
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.size = 4
    buffer1.bPut_Long Commands.CloseSession
    BlocksSocket1.Send buffer1.bytes
        
 
    If OkPress Then '[OK]
        If Plotf Then Exit Sub
        If SelectedResult > 0 And SelectedResult <= CountResults Then
            If Work_Result(SelectedResult).Vnutri_Figuri Then
                Razmestit_Reazultat Work_Result(SelectedResult), False, False, False, False, False, False, False, CheckBox1.value
            Else
                'Разместить результат
                Razmestit_Reazultat Work_Result(SelectedResult), Check4.value, Check5.value, Check3.value, Check6.value, Check7.value, Check8.value, Check9.value, CheckBox1.value
            End If
        End If
    End If
End Sub
