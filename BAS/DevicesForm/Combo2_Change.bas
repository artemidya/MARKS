Private Sub Combo2_Change()
    If Initialized Then
        With Profiles(PlotterForm.TabStrip1.SelectedItem.index)
            If MsgBox(localization.ActiveLng.Item("DevicesForm.MsgBox.ChangeLngCom"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then
                
                CommandButton3.Enabled = True
                .CommandLanguage = StringToCommandLanguage(Combo2.text)
                Select Case .CommandLanguage
                    Case CommandLanguages.HPGL
'                        Text4.text = localization.ActiveLng.Item("Profile.Device")
                        CheckBox7.Enabled = True
                        CheckBox7.value = .OneCommandOnPath
                    Case CommandLanguages.DMPL
'                        Text4.text = localization.ActiveLng.Item("Profile.DeviceDMPL")
                        CheckBox7.Enabled = False
                        CheckBox7.value = True
                End Select
                
                TextBox5.text = PlotterLanguage.Init
                TextBox8.text = PlotterLanguage.TERM
               
            Else
                Initialized = False
                Combo2.ListIndex = ListIndexByName(CommandLanguageToString(.CommandLanguage), Combo2)
                Initialized = True
            End If
        End With
    End If
End Sub
