Private Sub Command5_Click()
    'Cut
    If Combo1.ListIndex + 1 <> Combo1.ListCount Then
        'Activation Check
        ActivationCheck
        If Not PlotCalcIsActive Then Exit Sub
    End If
   
    'SendCutCommand
    SendCutCommand
    
    If PltDocument1(TabStrip1.value).RedLineError Then
        If MsgBox(localization.ActiveLng.Item("PlotterForm.MsgBox_1") & vbNewLine & localization.ActiveLng.Item("PlotterForm.MsgBox_2"), vbYesNoCancel Or vbCritical, "PlotCalc") <> vbYes Then Exit Sub
    End If

    '1. Converting to HPGL or DMPL
    Dim txt As String
    Dim NeedCut As Boolean
    Dim bytes() As Byte
    Dim f As Boolean
           
    If Profiles(TabStrip1.SelectedItem.index).CommandLanguage = HPGL Then
        Dim ConverterHPGL As New ConverterPlotCalcCrvToHPGL
        If RegData1.PreRun And Combo1.ListIndex + 1 <> Combo1.ListCount Then
            With PltDocument1(TabStrip1.value)
                txt = ConverterHPGL.ConvertPreRun(Profiles(TabStrip1.SelectedItem.index), OutPuts.x, OutPuts.y, .AllCurvesHeight, .AllCurvesWidth)
            End With
            bytes = StrToAscii(txt)
            OutPuts.Out bytes, OutPuts.x, OutPuts.y
            
            If MsgBox(localization.ActiveLng.Item("PlotterForm.MsgBox_3"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then NeedCut = True
        Else
            NeedCut = True
        End If

        If NeedCut Then
            With PltDocument1(TabStrip1.value)
                txt = ConverterHPGL.Convert(.Curves, Profiles(TabStrip1.SelectedItem.index), OutPuts.x, OutPuts.y, .AllCurvesHeight, .AllCurvesWidth)
                bytes = StrToAscii(txt)
                If Combo1.ListIndex + 1 <> Combo1.ListCount Then Me.Hide
                CreateAfterCutMark
                f = OutPuts.Out(bytes, OutPuts.x, OutPuts.y)
                If f Then
                    'UserForm_Terminate
                    Unload Me
                End If

                
            End With
        End If
    Else 'DMPL
        Dim ConverterDMPL As New ConverterPlotCalcCrvToDMPL
        If RegData1.PreRun And Combo1.ListIndex + 1 <> Combo1.ListCount Then
            With PltDocument1(TabStrip1.value)
                txt = ConverterDMPL.ConvertPreRun(Profiles(TabStrip1.SelectedItem.index), OutPuts.x, OutPuts.y, .AllCurvesHeight, .AllCurvesWidth)
            End With
            bytes = StrToAscii(txt)
            OutPuts.Out bytes, OutPuts.x, OutPuts.y
            If MsgBox(localization.ActiveLng.Item("PlotterForm.MsgBox_3"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then NeedCut = True
        Else
            NeedCut = True
        End If
        
        If NeedCut Then
            With PltDocument1(TabStrip1.value)
                txt = ConverterDMPL.Convert(.Curves, Profiles(TabStrip1.SelectedItem.index), OutPuts.x, OutPuts.y, .AllCurvesHeight, .AllCurvesWidth)
                bytes = StrToAscii(txt)
                If Combo1.ListIndex + 1 <> Combo1.ListCount Then Me.Hide
                CreateAfterCutMark
                f = OutPuts.Out(bytes, OutPuts.x, OutPuts.y)
                If f Then
                    'UserForm_Terminate
                    Unload Me
                End If

            End With
        End If
    End If
End Sub
