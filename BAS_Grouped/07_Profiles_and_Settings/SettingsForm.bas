' =============================================================================
' Group: 07_Profiles_and_Settings
' Source module: SettingsForm.frm
' Procedures: 50
' =============================================================================

' -----------------------------------------------------------------------------
' CheckBox1_Change (Sub), source lines 29-33
' -----------------------------------------------------------------------------
Private Sub CheckBox1_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox2_Change (Sub), source lines 34-38
' -----------------------------------------------------------------------------
Private Sub CheckBox2_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox3_Change (Sub), source lines 39-43
' -----------------------------------------------------------------------------
Private Sub CheckBox3_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox4_Change (Sub), source lines 44-48
' -----------------------------------------------------------------------------
Private Sub CheckBox4_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox5_Change (Sub), source lines 50-62
' -----------------------------------------------------------------------------
Private Sub CheckBox5_Change()
    If CheckBox5.value Then
        Label13.Enabled = False
        TextBox3.Enabled = False
    Else
        Label13.Enabled = True
        TextBox3.Enabled = True
    End If

    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox6_Change (Sub), source lines 65-69
' -----------------------------------------------------------------------------
Private Sub CheckBox6_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox7_Change (Sub), source lines 72-76
' -----------------------------------------------------------------------------
Private Sub CheckBox7_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox1_Change (Sub), source lines 80-84
' -----------------------------------------------------------------------------
Private Sub TextBox1_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox2_Change (Sub), source lines 85-89
' -----------------------------------------------------------------------------
Private Sub TextBox2_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox3_Change (Sub), source lines 91-95
' -----------------------------------------------------------------------------
Private Sub TextBox3_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 97-135
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    UserData.Load
    
    Combo2.AddItem
    Combo2.ListIndex = 0
    SetTexts

    IsEditingMyFields = True
    
    Combo1.List = Array("RU", "EN", "UK", "ZH", "PL", "DE", "FR", "SK", "SL", "IT", "VI", "KK", "PT", "ES", "TR", "ID", "AR", "HI")
    SelectLangInCombo (UserData.LangCode)
    
    Text4.text = Trim(UserData.PerimetrFormData.MoneyUnit)
    Label4.Caption = UCase(Trim(UserData.CodeOfActivation))

    Dim version As String
    version = left(UserData.PlotCalcVersion, 4) & "." & Mid(UserData.PlotCalcVersion, 5, 2) & "." & right(UserData.PlotCalcVersion, 2)
    Label5.Caption = version
    CheckBox1.value = RegData1.AutoUpdate
    If Len(Trim(UserData.CodeOfActivation)) > 0 Then
        CommandButton2.Enabled = True
        Label6.Enabled = True
    End If
    
    CheckBox2.value = RegData1.DelFill
    CheckBox3.value = RegData1.DelOutline
    CheckBox4.value = RegData1.UngroupAll
    CheckBox5.value = RegData1.OutlineUse
    CheckBox6.value = RegData1.ConnectingCurves
    CheckBox7.value = RegData1.ReduceCurves
    
        
    TextBox1.text = CDbl(RegData1.CropBMPoffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox2.text = CDbl(RegData1.TrimOffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox3.text = CDbl(RegData1.Sensitivity) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm")

    Set UpdateCenter1 = New UpdateCenter
    IsEditingMyFields = False
End Sub

' -----------------------------------------------------------------------------
' CommandButton3_Click (Sub), source lines 136-172
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' UpdateCenter1_DownloadPercent (Sub), source lines 174-176
' -----------------------------------------------------------------------------
Private Sub UpdateCenter1_DownloadPercent(Percent As Long)
    Label9.Caption = Percent & "%"
End Sub

' -----------------------------------------------------------------------------
' UpdateCenter1_Error (Sub), source lines 177-181
' -----------------------------------------------------------------------------
Private Sub UpdateCenter1_Error()
    Err = True
    Label9.Caption = "error"
    CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' RefreshUpdateButton (Sub), source lines 182-189
' -----------------------------------------------------------------------------
Private Sub RefreshUpdateButton()
    If UpdateStatus = GetVersion Then
        CommandButton3.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton3.CaptionCheck")
    Else
        CommandButton3.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton3.CaptionRefresh")
    End If
'
End Sub

' -----------------------------------------------------------------------------
' Combo1_MouseMove (Sub), source lines 190-193
' -----------------------------------------------------------------------------
Private Sub Combo1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Combo2_MouseMove (Sub), source lines 194-197
' -----------------------------------------------------------------------------
Private Sub Combo2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Command5_MouseMove (Sub), source lines 198-201
' -----------------------------------------------------------------------------
Private Sub Command5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' CommandButton1_MouseMove (Sub), source lines 202-205
' -----------------------------------------------------------------------------
Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' CommandButton2_MouseMove (Sub), source lines 206-209
' -----------------------------------------------------------------------------
Private Sub CommandButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Frame1_MouseMove (Sub), source lines 210-213
' -----------------------------------------------------------------------------
Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Frame3_MouseMove (Sub), source lines 214-217
' -----------------------------------------------------------------------------
Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Label1_MouseMove (Sub), source lines 218-221
' -----------------------------------------------------------------------------
Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Label2_MouseMove (Sub), source lines 222-225
' -----------------------------------------------------------------------------
Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Label3_MouseMove (Sub), source lines 226-229
' -----------------------------------------------------------------------------
Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Label4_MouseMove (Sub), source lines 230-233
' -----------------------------------------------------------------------------
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Label5_MouseMove (Sub), source lines 234-237
' -----------------------------------------------------------------------------
Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Text4_MouseMove (Sub), source lines 238-241
' -----------------------------------------------------------------------------
Private Sub Text4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' Frame2_MouseMove (Sub), source lines 242-245
' -----------------------------------------------------------------------------
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' UserForm_MouseMove (Sub), source lines 246-249
' -----------------------------------------------------------------------------
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.Font.Underline = True
    DoEvents
End Sub

' -----------------------------------------------------------------------------
' SetTexts (Sub), source lines 250-297
' -----------------------------------------------------------------------------
Private Sub SetTexts()
    Me.Caption = localization.ActiveLng.Item("SettingsForm.Caption")
    Label6.Caption = localization.ActiveLng.Item("SettingsForm.Label6.Caption")
    Label6.left = 91 - Label6.Width \ 2
    Label6.ControlTipText = localization.ActiveLng.Item("SettingsForm.TestCodeLink") & UserData.CodeOfActivation
    
    Frame1.Caption = localization.ActiveLng.Item("SettingsForm.Frame1.Caption")
    Label1.Caption = localization.ActiveLng.Item("SettingsForm.Label1.Caption")
    Label2.Caption = localization.ActiveLng.Item("SettingsForm.Label2.Caption")
    Label3.Caption = localization.ActiveLng.Item("SettingsForm.Label3.Caption")
    Frame3.Caption = localization.ActiveLng.Item("SettingsForm.Frame3.Caption")
    Frame2.Caption = localization.ActiveLng.Item("SettingsForm.Frame2.Caption")
    CommandButton2.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton2.Caption")
    CommandButton1.Caption = localization.ActiveLng.Item("SettingsForm.CommandButton1.Caption")
    Command5.Caption = localization.ActiveLng.Item("SettingsForm.Command5.Caption")
    Combo2.List(0) = localization.ActiveLng.Item("SettingsForm.Combo2.Item1")
    
    Label7.Caption = localization.ActiveLng.Item("SettingsForm.Label7")
    Label8.Caption = localization.ActiveLng.Item("SettingsForm.Label8")
    CheckBox1.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox1.Caption")
    
    Frame4.Caption = localization.ActiveLng.Item("SettingsForm.AdditionalParameters")
    Label10.Caption = localization.ActiveLng.Item("SettingsForm.OffsetCropping")
    Label11.Caption = localization.ActiveLng.Item("SettingsForm.OffsetTrimming")
    
    Label12.Caption = localization.ActiveLng.Item("SettingsForm.Label12")
    CheckBox2.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox2")
    CheckBox3.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox3")
    CheckBox4.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox4")

    Label13.Caption = localization.ActiveLng.Item("SettingsForm.Label13")
    CheckBox5.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox5")
    
    CheckBox6.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox6")
    CheckBox7.Caption = localization.ActiveLng.Item("SettingsForm.CheckBox7")
    
    

    Dim t As Boolean
    t = IsEditingMyFields
    IsEditingMyFields = True
    TextBox1.text = ValP(TextBox1.text) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox2.text = ValP(TextBox2.text) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox3.text = ValP(TextBox3.text) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    IsEditingMyFields = t
    
    RefreshUpdateButton
End Sub

' -----------------------------------------------------------------------------
' SetAllTexts (Sub), source lines 298-307
' -----------------------------------------------------------------------------
Private Sub SetAllTexts()
    If CalculateSanP.PerimetrFormVisible Then
        PerimetrForm.UserFormRefreshData
        PerimetrForm.SetTexts
    End If
    If CalculateSanP.SquareFormVisible Then
        SquareForm.UserFormRefreshData
        SquareForm.SetTexts
    End If
End Sub

' -----------------------------------------------------------------------------
' Combo1_Change (Sub), source lines 308-315
' -----------------------------------------------------------------------------
Private Sub Combo1_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
    SelectLangByCode (Combo1.value)
    SetTexts
    SetAllTexts
End Sub

' -----------------------------------------------------------------------------
' Combo1_KeyDown (Sub), source lines 316-318
' -----------------------------------------------------------------------------
Private Sub Combo1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' Combo2_KeyDown (Sub), source lines 319-321
' -----------------------------------------------------------------------------
Private Sub Combo2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' Command5_KeyDown (Sub), source lines 322-324
' -----------------------------------------------------------------------------
Private Sub Command5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' CommandButton1_KeyDown (Sub), source lines 325-327
' -----------------------------------------------------------------------------
Private Sub CommandButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' CommandButton2_Click (Sub), source lines 328-338
' -----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    If MsgBox(localization.ActiveLng.Item("SettingsForm.QuestionDeleteThisCode"), vbYesNo Or vbQuestion, "PlotCalc") = vbYes Then
        UserData.CodeOfActivation = ""
        UserData.Save
        Label4.Caption = ""
        CommandButton2.Enabled = False
        Label6.Enabled = False
        CommandButton1.Enabled = True
        CommandButton1.SetFocus
    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton2_KeyDown (Sub), source lines 339-341
' -----------------------------------------------------------------------------
Private Sub CommandButton2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' Label6_Click (Sub), source lines 342-346
' -----------------------------------------------------------------------------
Private Sub Label6_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", localization.ActiveLng.Item("SettingsForm.TestCodeLink") & Trim(UserData.CodeOfActivation), "", "", 4
10:
End Sub

' -----------------------------------------------------------------------------
' Label6_MouseMove (Sub), source lines 347-353
' -----------------------------------------------------------------------------
Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Label6.Font.Underline Then
        Label6.Font.Underline = False
        DoEvents
    End If
    Cursor.SetCursor2 Cursor.IDC_HAND
End Sub

' -----------------------------------------------------------------------------
' Text4_Change (Sub), source lines 354-358
' -----------------------------------------------------------------------------
Private Sub Text4_Change()
    If IsEditingMyFields Then Exit Sub
    DataHasBeenChanged = True
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' Text4_KeyDown (Sub), source lines 359-361
' -----------------------------------------------------------------------------
Private Sub Text4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' SelectLangInCombo (Sub), source lines 362-368
' -----------------------------------------------------------------------------
Private Sub SelectLangInCombo(Lng As String)
    On Error GoTo 10
    Combo1.value = UCase(Lng)
    Exit Sub
10:
    Combo1.value = "EN"
End Sub

' -----------------------------------------------------------------------------
' Command5_Click (Sub), source lines 369-371
' -----------------------------------------------------------------------------
Private Sub Command5_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' CommandButton1_Click (Sub), source lines 372-376
' -----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    SaveUserData
    DataHasBeenChanged = False
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' UserForm_QueryClose (Sub), source lines 377-392
' -----------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If DataHasBeenChanged Then
        Dim res As VbMsgBoxResult
        res = MsgBox(localization.ActiveLng.Item("SettingsForm.QuestionSaveChenges"), vbYesNoCancel Or vbQuestion, "PlotCalc")
        Select Case res
            Case vbYes
                SaveUserData
            Case vbNo
                'Exit without saving
                SelectLangByCode (UserData.LangCode)
                SetAllTexts
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub

' -----------------------------------------------------------------------------
' ValP (Function), source lines 393-395
' -----------------------------------------------------------------------------
Private Function ValP(ByRef txt As String) As Double
    ValP = Abs(val(Replace$(txt, ",", ".")))
End Function

' -----------------------------------------------------------------------------
' LngP (Function), source lines 396-400
' -----------------------------------------------------------------------------
Private Function LngP(val As Double) As Long
    On Error GoTo 10
    LngP = CLng(val)
10
End Function

' -----------------------------------------------------------------------------
' SaveUserData (Sub), source lines 401-421
' -----------------------------------------------------------------------------
Private Sub SaveUserData()
    'Save changes....
    RegData1.AutoUpdate = CheckBox1.value
    RegData1.CropBMPoffset = LngP(ValP(TextBox1.text) * 1000)
    RegData1.TrimOffset = LngP(ValP(TextBox2.text) * 1000)
    RegData1.Sensitivity = LngP(ValP(TextBox3.text) * 1000)
    
    RegData1.DelFill = CheckBox2.value
    RegData1.DelOutline = CheckBox3.value
    RegData1.UngroupAll = CheckBox4.value
    
    RegData1.OutlineUse = CheckBox5.value
    RegData1.ConnectingCurves = CheckBox6.value
    RegData1.ReduceCurves = CheckBox7.value
    
    UserData.PerimetrFormData.MoneyUnit = Text4.text
    UserData.LangCode = Combo1.value
    ToolBar.SetToolTipText
    UserData.Save
    SetAllTexts
End Sub
