Private Sub BlocksSocket1_OnRecive(bytes() As Byte, client As BlocksSocket)
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.SetBytes bytes
    Dim command As Commands
    command = buffer1.bGet_Long
    Dim content As String
    Select Case command
        Case Commands.SetPlotterList
            content = buffer1.bGet_String((UBound(bytes) - 3) \ 2)
            PlottersList() = GetPlottersList(content)
            NeedAddList() = GetNeedAddList(content)
            Label40.Caption = localization.ActiveLng.Item("DevicesForm.Label40.Caption_LastUpdt") & " " & DateLngToString(FindMaxDate(PlottersList))
            RefreshLists
            Dim ForModeratorDevice As String
            ForModeratorDevice = Profiles(PlotterForm.TabStrip1.SelectedItem.index).ForModeratorDevice
            If ForModeratorDevice <> "" Then
                Dim q As Long
                If Not Not PlottersList Then
                    If UBound(PlottersList) > -1 Then
                        For q = 0 To UBound(PlottersList)
                            If PlottersList(q).Device = ForModeratorDevice Then
                                MsgBox localization.ActiveLng.Item("DevicesForm.InputBox") & " (" & ForModeratorDevice & ")", vbInformation, "PlotCalc"
                                Profiles(PlotterForm.TabStrip1.SelectedItem.index).ForModeratorDevice = ""
                                GoTo ok
                            End If
                        Next
                    End If
                End If
                If Not Not NeedAddList Then
                    If UBound(NeedAddList) > -1 Then
                        For q = 0 To UBound(NeedAddList)
                            If NeedAddList(q).Device = ForModeratorDevice Then
                                GoTo ok
                            End If
                        Next
                    End If
                End If
                MsgBox localization.ActiveLng.Item("DevicesForm.MsgBox_M2") & " (" & ForModeratorDevice & ")", vbInformation, "PlotCalc"
                Profiles(PlotterForm.TabStrip1.SelectedItem.index).ForModeratorDevice = ""
ok:
            End If

        Case Commands.SetPlotter
            content = buffer1.bGet_String((UBound(bytes) - 3) \ 2)
            Dim testVersion As Long
            testVersion = StringToLong(GetTag("MinimalVersion", content))
            If testVersion > UserData.PlotCalcVersion Then
                MsgBox localization.ActiveLng.Item("Profile.MsgBox_1") & vbNewLine & vbNewLine & localization.ActiveLng.Item("Profile.MsgBox_2") & vbNewLine & localization.ActiveLng.Item("Profile.MsgBox_3") & vbNewLine & localization.ActiveLng.Item("Profile.MsgBox_4"), vbInformation
            Else
                If MsgBox(localization.ActiveLng.Item("DevicesForm.MsgBox.Repl"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then
                    InsertContent content
                    Frame4.ScrollTop = 0
                End If
            End If
            LoadPlotterProcess = False
            Label40.Caption = TempLabel40
        Case Commands.PlotterAddedSuccessfully
            MsgBox localization.ActiveLng.Item("DevicesForm.MsgBox.willadded") & vbNewLine & localization.ActiveLng.Item("DevicesForm.MsgBox.info"), vbInformation, "PlotCalc"
        Case Commands.ErrorWrongPassword
            Authorization True
    End Select
End Sub
