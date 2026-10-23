' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: PlotterForm.frm
' Procedures: 97
' =============================================================================

' -----------------------------------------------------------------------------
' AddOrDeletItemsInCombo (Sub), source lines 63-74
' -----------------------------------------------------------------------------
Private Sub AddOrDeletItemsInCombo(n As Long, Combo As ComboBox)
    Dim q As Long
    If n < 0 Then
        For q = 1 To Abs(n)
            Combo.RemoveItem (Combo.ListCount - 1)
        Next
    Else
        For q = 1 To n
            Combo.AddItem ""
        Next
    End If
End Sub

' -----------------------------------------------------------------------------
' OutPuts_WiFiChange (Sub), source lines 82-105
' -----------------------------------------------------------------------------
Private Sub OutPuts_WiFiChange()
    Dim q As Long
    q = Combo1.ListCount - 2

    With Combo1
       
        If OutPuts.Connected(q) = True Then
            If .List(q) <> ChrW(10003) & OutPuts.Names(q) Then
                suppressChange = True
                .List(q) = ChrW(10003) & OutPuts.Names(q)
                suppressChange = False
                Combo1_Change
            End If
        Else
            If .List(q) <> "    " & OutPuts.Names(q) Then
                suppressChange = True
                .List(q) = "    " & OutPuts.Names(q)
                suppressChange = False
                Combo1_Change
            End If
        End If
       
    End With
End Sub

' -----------------------------------------------------------------------------
' USB_Events1_Change (Sub), source lines 109-143
' -----------------------------------------------------------------------------
Private Sub USB_Events1_Change()
    Dim selItem As String
    Dim selIndex As Long
    
    OutPuts.Refresh
    
    With Combo1
        selItem = .text
        selIndex = .ListIndex
        suppressChange = True
    
        AddOrDeletItemsInCombo OutPuts.count - .ListCount, Combo1
        For q = 0 To Combo1.ListCount - 1
            If OutPuts.Connected(q) = True Then
                If .List(q) <> ChrW(10003) & OutPuts.Names(q) Then
                    .List(q) = ChrW(10003) & OutPuts.Names(q)
                End If
            Else
                If .List(q) <> "    " & OutPuts.Names(q) Then
                    .List(q) = "    " & OutPuts.Names(q)
                End If
            End If
        Next
        suppressChange = False
        If .text <> selItem Then
            .ListIndex = ListIndexByName(KillBirdsAndSpaces(selItem), Combo1, selIndex)
        End If
    End With
    Combo1_Change
    
    If COMSettingsForm.isActive Then
        Unload COMSettingsForm
        OutPuts.Configure KillBirdsAndSpaces(Combo1.text), FindWindow(vbNullString, Me.Caption)
    End If
End Sub

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 145-256
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    PlotCalcIsActive = True
    UserData.Load
    Set OutPuts = New OutPuts
    Deployed = RegData1.Deployed
    AddLog "'help' or '?' - commands help", True
    ReDim Profiles(7)
    Set Profiles(0) = New Profile
    Set Profiles(1) = New Profile
    Set Profiles(2) = New Profile
    Set Profiles(3) = New Profile
    Set Profiles(4) = New Profile
    Set Profiles(5) = New Profile
    Set Profiles(6) = New Profile
    Set Profiles(7) = New Profile
    Set Converter = New ConverterCorelCrvToPlotCalcCrv
    
    TextBox1.text = Profiles(0).name
    TextBox2.text = Profiles(1).name
    TextBox3.text = Profiles(2).name
    TextBox4.text = Profiles(3).name
    TextBox5.text = Profiles(4).name
    TextBox6.text = Profiles(5).name
    TextBox7.text = Profiles(6).name
    TextBox8.text = Profiles(7).name
    
    CheckBox1.value = Not Profiles(0).Visible
    CheckBox2.value = Not Profiles(1).Visible
    CheckBox3.value = Not Profiles(2).Visible
    CheckBox4.value = Not Profiles(3).Visible
    CheckBox5.value = Not Profiles(4).Visible
    CheckBox6.value = Not Profiles(5).Visible
    CheckBox7.value = Not Profiles(6).Visible
    CheckBox8.value = Not Profiles(7).Visible

    ListenKeyboard = False
    Dim q As Long
    With Combo1
        For q = 0 To OutPuts.count - 1
            If OutPuts.Connected(q) = True Then
                .AddItem ChrW(10003) & OutPuts.Names(q)
            Else
                .AddItem "    " & OutPuts.Names(q)
            End If
        Next
    End With
    
    If RegData1.ActiveProfile < TabStrip1.Tabs.count Then TabStrip1.value = RegData1.ActiveProfile
    RefreshFields
    
    Initialized = True
    
    
    
    'Captions
    
    Frame3.Caption = localization.ActiveLng.Item("PlotterForm.Frame3.Caption")  '"Hardware"
    Label9.Caption = localization.ActiveLng.Item("PlotterForm.Label9.Caption")  '"Device"
    Label10.Caption = localization.ActiveLng.Item("PlotterForm.Label10.Caption")  '"Connection"
    CommandButton22.Caption = localization.ActiveLng.Item("PlotterForm.CommandButton22.Caption")  '"Test"
    Frame4.Caption = localization.ActiveLng.Item("PlotterForm.Frame4.Caption")  '"Parametrs"
    Label14.Caption = localization.ActiveLng.Item("PlotterForm.Label14.Caption")  '"Acceleration"
    Label12.Caption = localization.ActiveLng.Item("PlotterForm.Label12.Caption")  '"Velocity"
    Label12.ControlTipText = localization.ActiveLng.Item("PlotterForm.Label12.ControlTipText")  '"Double click changes options"
    ImageSpeed.ControlTipText = Label12.ControlTipText
    Label13.Caption = localization.ActiveLng.Item("PlotterForm.Label13.Caption")  '"Force"
    Frame5.Caption = localization.ActiveLng.Item("PlotterForm.Frame5.Caption")  '"Curve parameters"
    Label15.Caption = localization.ActiveLng.Item("PlotterForm.Label15.Caption")  '"Knife offset"
    TextBox10.ControlTipText = "(0-0,500 " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    Label11.Caption = localization.ActiveLng.Item("PlotterForm.Label11.Caption")  '"Closing curve"
    TextBox9.ControlTipText = "(0-10,000 " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    CheckBox10.Caption = ChrW(215) & "2"
    CheckBox10.ControlTipText = localization.ActiveLng.Item("PlotterForm.CheckBox10.ControlTipText")  '"Double-tracking. Each path will be passed through twice"
    Frame7.Caption = localization.ActiveLng.Item("PlotterForm.Frame7.Caption")  '"Crop Marks"
    Frame6.Caption = localization.ActiveLng.Item("PlotterForm.Frame6.Caption")  '"Margin"
    CommandButton23.ControlTipText = localization.ActiveLng.Item("PlotterForm.CommandButton23.ControlTipText")  '"On/off keypad control"
    CheckBox9.Caption = localization.ActiveLng.Item("PlotterForm.CheckBox9.Caption")  '"Pre-run round"
    CommandButton24.Caption = localization.ActiveLng.Item("PlotterForm.CommandButton24.Caption")  '"Show Way"
    Label20.Caption = localization.ActiveLng.Item("PlotterForm.Label20.Caption")  '"Command line"
    CommandButton20.Caption = localization.ActiveLng.Item("PlotterForm.CommandButton20.Caption")  '"ENTER"
    Frame1.Caption = localization.ActiveLng.Item("PlotterForm.Frame1.Caption")  '"Settings"
    Label1.Caption = localization.ActiveLng.Item("PlotterForm.Label1.Caption")  '"Name profile 1"
    Label2.Caption = Label1.Caption & " 2"
    Label3.Caption = Label1.Caption & " 3"
    Label4.Caption = Label1.Caption & " 4"
    Label5.Caption = Label1.Caption & " 5"
    Label6.Caption = Label1.Caption & " 6"
    Label7.Caption = Label1.Caption & " 7"
    Label8.Caption = Label1.Caption & " 8"
    Label1.Caption = Label1.Caption & " 1"
    
    CheckBox1.Caption = localization.ActiveLng.Item("PlotterForm.CheckBox1.Caption")  '"Hide"
    CheckBox1.Width = 64.5
    CheckBox2.Caption = CheckBox1.Caption
    CheckBox3.Caption = CheckBox1.Caption
    CheckBox4.Caption = CheckBox1.Caption
    CheckBox5.Caption = CheckBox1.Caption
    CheckBox6.Caption = CheckBox1.Caption
    CheckBox7.Caption = CheckBox1.Caption
    CheckBox8.Caption = CheckBox1.Caption
    CommandButton3.Caption = localization.ActiveLng.Item("PlotterForm.CommandButton3.Caption")  '"Show file"
    CommandButton4.Caption = CommandButton3.Caption
    CommandButton5.Caption = CommandButton3.Caption
    CommandButton6.Caption = CommandButton3.Caption
    CommandButton7.Caption = CommandButton3.Caption
    CommandButton8.Caption = CommandButton3.Caption
    CommandButton9.Caption = CommandButton3.Caption
    CommandButton10.Caption = CommandButton3.Caption
    CommandButton21.ControlTipText = localization.ActiveLng.Item("PlotterForm.CommandButton21.ControlTipText")
    CommandButton22.ControlTipText = localization.ActiveLng.Item("PlotterForm.CommandButton22.ControlTipText") 'Проверка подключения (после нажатия каретка плоттера должна сделать небольшое движение)
    Label22.Caption = localization.ActiveLng.Item("PlotterForm.Label22.Caption")
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 257-277
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()
    If Not UserForm_Activ Then
        UserForm_Activ = True
        'Read data from CorelDraw
        Set OriginalCurves = Converter.GetCurves(OriginalHeight, OriginalWidth, MaxColor)
        
        'Init BlocksSokets
        BlocksSocketShared.HookForm FindWindow(vbNullString, Me.Caption)
        Set BlocksSocket1 = New BlocksSocket
        BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
        Set USB_Events1 = New USB_Events
        
        
        If OriginalCurves.count = 0 Then Command5.Enabled = False
        Label21.Visible = False
        If OriginalCurves.count > 0 Then CommandButton24.Enabled = True
        RefreshPltDocument
        
        KeyReadLoop
    End If
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnConnect (Sub), source lines 278-315
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnConnect(client As BlocksSocket)
    'Send Head
    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 64
        .bPut_Long Commands.HEAD_CutWindow_20250528 '4
        .bPut_Long UserData.PlotCalcVersion         '4
        .bPut_Long RegData1.FirstCutDateLng         '4
'        .bPut_Long 20250510
        .bPut_Long OriginalCurves.count             '4


        If Application.VersionMajor = 24 And Application.VersionMinor >= 3 Then
            'Corel 2023
            .bPut_Long Application.VersionMajor + 1999 '4
        Else
            If Application.VersionMajor >= 25 Then
                'Corel 2024 +
                .bPut_Long Application.VersionMajor + 1999 '4
            Else
                .bPut_Long Application.VersionMajor + 1998 '4
            End If
        End If
        

        Dim lcode As String * 2
        lcode = UserData.LangCode
        .bPut_String lcode '4
        Dim code As String * 20
        code = Trim(UserData.CodeOfActivation)
        If Len(code) <> 20 Then code = "--------------------"
        .bPut_String code '40

        BlocksSocket1.Send .bytes
        
    End With
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnRecive (Sub), source lines 316-352
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' BlocksSocket1_OnDisconnect (Sub), source lines 353-358
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnDisconnect(client As BlocksSocket)
    If ActivationFormIsShow Then
        ActivationForm.Command1.Enabled = False
        ActivationForm.Text1.Enabled = False
    End If
End Sub

' -----------------------------------------------------------------------------
' ImageSpeed_Click (Sub), source lines 359-362
' -----------------------------------------------------------------------------
Private Sub ImageSpeed_Click()
    Profiles(TabStrip1.SelectedItem.index).SeparateParametrs = Not Profiles(TabStrip1.SelectedItem.index).SeparateParametrs
    RefreshFields
End Sub

' -----------------------------------------------------------------------------
' Label12_Click (Sub), source lines 364-367
' -----------------------------------------------------------------------------
Private Sub Label12_Click()
    Profiles(TabStrip1.SelectedItem.index).SeparateParametrs = Not Profiles(TabStrip1.SelectedItem.index).SeparateParametrs
    RefreshFields
End Sub

' -----------------------------------------------------------------------------
' CheckBox10_Change (Sub), source lines 368-373
' -----------------------------------------------------------------------------
Private Sub CheckBox10_Change()
    Profiles(TabStrip1.SelectedItem.index).DoublePass = CheckBox10.value
    CommandButton18.Enabled = Not CheckBox10.value
    CommandButton19.Enabled = Not CheckBox10.value
    TextBox9.Enabled = Not CheckBox10.value
End Sub

' -----------------------------------------------------------------------------
' ActivationCheck (Sub), source lines 374-395
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' DateToLong (Function), source lines 396-398
' -----------------------------------------------------------------------------
Private Function DateToLong(myDate As Date) As Long
      DateToLong = Year(myDate) * 10000 + Month(myDate) * 100 + Day(myDate)
End Function

' -----------------------------------------------------------------------------
' SendCutCommand (Sub), source lines 399-431
' -----------------------------------------------------------------------------
Private Sub SendCutCommand()
    Dim cp As Profile
    Set cp = Profiles(TabStrip1.SelectedItem.index)
    Dim ComSettingsStr As String
    ComSettingsStr = OutPuts.ComSettingsStr

    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 52 + LenB(Text4.text) + LenB(KillBirdsAndSpaces(Combo1.text)) + LenB(OutPuts.VidAndPid(KillBirdsAndSpaces(Combo1.text))) + LenB(OutPuts.ExparedID(KillBirdsAndSpaces(Combo1.text))) + LenB(ComSettingsStr)
        .bPut_Long Commands.Cut
        .bPut_Long cp.CommandLanguage
        .bPut_Long cp.ZeroPosition
        .bPut_Long cp.CoordinateSystem
        .bPut_Long cp.LineType
        .bPut_Boolean cp.OneCommandOnPath
        .bPut_Boolean False 'plotter is autoselected
        .bPut_Long Len(Text4.text)
        .bPut_String Text4.text
        .bPut_Long Len(KillBirdsAndSpaces(Combo1.text))
        .bPut_String KillBirdsAndSpaces(Combo1.text)
        .bPut_Long Len(OutPuts.VidAndPid(KillBirdsAndSpaces(Combo1.text)))
        .bPut_String OutPuts.VidAndPid(KillBirdsAndSpaces(Combo1.text))
        .bPut_Long Len(OutPuts.ExparedID(KillBirdsAndSpaces(Combo1.text)))
        .bPut_String OutPuts.ExparedID(KillBirdsAndSpaces(Combo1.text))
        .bPut_Long CLng(cp.WidthScale * 10000)
        .bPut_Long CLng(cp.HeightScale * 10000)
        .bPut_Long Len(ComSettingsStr)
        .bPut_String ComSettingsStr
        
        BlocksSocket1.Send .bytes
    End With
End Sub

' -----------------------------------------------------------------------------
' Command5_Click (Sub), source lines 433-511
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' CreateAfterCutMark (Sub), source lines 512-550
' -----------------------------------------------------------------------------
Private Sub CreateAfterCutMark()
    Application.ActiveDocument.BeginCommandGroup "Mark"
    '\ /
    ' X
    '/ \
    Dim crv As VGCore.curve
    Dim P1 As waGraphicsShared.POINTF
    Dim P2 As waGraphicsShared.POINTF
    Dim P3 As waGraphicsShared.POINTF
    Dim P4 As waGraphicsShared.POINTF
    Dim d As Double:  d = 0.5
    
    With VGCore.ActiveSelection
        P1.x = .LeftX - d
        P1.y = .BottomY - d
        P2.x = .RightX + d
        P2.y = .TopY + d
        
        P3.x = .LeftX - d
        P3.y = .TopY + d
        P4.x = .RightX + d
        P4.y = .BottomY - d
    End With
    
    Set crv = VGCore.ActiveDocument.CreateCurve
    With crv.CreateSubPath(P1.x, P1.y)
        .AppendLineSegment P2.x, P2.y
    End With

    With crv.CreateSubPath(P3.x, P3.y)
        .AppendLineSegment P4.x, P4.y
    End With
    
    Dim s1 As VGCore.shape
    Set s1 = VGCore.ActiveLayer.CreateCurve(crv)
    s1.Outline.SetProperties Width:=0.1, Color:=CreateCMYKColor(0, 100, 100, 0)
    
    Application.ActiveDocument.EndCommandGroup
End Sub

' -----------------------------------------------------------------------------
' CommandButton24_Click (Sub), source lines 551-554
' -----------------------------------------------------------------------------
Private Sub CommandButton24_Click()
    PltDocument1(TabStrip1.value).ShowWay
    
End Sub

' -----------------------------------------------------------------------------
' StopShowAll (Sub), source lines 555-562
' -----------------------------------------------------------------------------
Private Sub StopShowAll()
    Dim q As Long
    For q = 0 To 7
        If Not PltDocument1(q) Is Nothing Then
            PltDocument1(q).StopShowWay
        End If
    Next
End Sub

' -----------------------------------------------------------------------------
' RefreshPltDocument (Sub), source lines 564-625
' -----------------------------------------------------------------------------
Private Sub RefreshPltDocument()
    If TabStrip1.value > 7 Then Exit Sub
    If OriginalCurves Is Nothing Then Exit Sub
    If PltDocument1(TabStrip1.value) Is Nothing Then
        Me.Repaint
        Set PltDocument1(TabStrip1.value) = New PltDocument
        With PltDocument1(TabStrip1.value)
            Dim kdpi As Double
            kdpi = MonitorDPI.GetDpiDbl
            .SetScreenSize 760 * kdpi, 552 * kdpi
            .PositionX = 400 * kdpi
            .PositionY = 47 * kdpi
            
            If Profiles(TabStrip1.value).MaxSheetHeight = 0 Then
                .MaxSheetHeight = 2147483647
            Else
                .MaxSheetHeight = Profiles(TabStrip1.value).MaxSheetHeight * 1000
            End If
            
            If Profiles(TabStrip1.value).MaxSheetWidth = 0 Then
                If Profiles(TabStrip1.value).Turn = Turn0 Or Profiles(TabStrip1.value).Turn = Turn180 Then
                    If OriginalWidth <= 600000 Then
                        .MaxSheetWidth = 600000
                    ElseIf OriginalWidth <= 1600000 Then
                        .MaxSheetWidth = 1600000
                    Else
                        .MaxSheetWidth = OriginalWidth
                    End If
                Else
                    If OriginalHeight <= 600000 Then
                        .MaxSheetWidth = 600000
                    ElseIf OriginalHeight <= 1600000 Then
                        .MaxSheetWidth = 1600000
                    Else
                        .MaxSheetWidth = OriginalHeight
                    End If
                End If
            Else
                .MaxSheetWidth = Profiles(TabStrip1.value).MaxSheetWidth * 1000
            End If
          

            .PlotterModelString = Profiles(TabStrip1.value).PlotterModelString
            .PlotterName = Profiles(TabStrip1.value).Device

            

            Set .MaterialBrush = CreateBrush(RGB(MaxColor.R, MaxColor.g, MaxColor.b), CSng(MaxColor.a) / 255 * 100)
            .SetCurvesByVal OriginalCurves, OriginalHeight, OriginalWidth, Profiles(TabStrip1.value).Turn, Profiles(TabStrip1.value).HeightMirror, Profiles(TabStrip1.value).WidthMirror
            .ZeroPosition = Profiles(TabStrip1.value).ZeroPosition
            
            Dim p As POINTL
            p.x = OutPuts.Xmm * 1000
            p.y = OutPuts.Ymm * 1000
            PltDocument1(TabStrip1.value).XYmargin = p
            
            .Render
        End With
    Else
        PltDocument1(TabStrip1.value).Refresh
    End If
End Sub

' -----------------------------------------------------------------------------
' Converter_OnPercentChange (Sub), source lines 626-628
' -----------------------------------------------------------------------------
Private Sub Converter_OnPercentChange(Percent As Long)
        Label21.Caption = Percent & "%"
End Sub

' -----------------------------------------------------------------------------
' ListenKeyboard (PropertyLet), source lines 629-640
' -----------------------------------------------------------------------------
Private Property Let ListenKeyboard(oValue As Boolean)
    lListenKeyboard = oValue
    If oValue Then
        'Reset buffer
        Dim q As Long
        For q = 0 To 12 'Count reading keys
            PressedKey
        Next
    Else
        CommandButton23.Picture = Image3.Picture
    End If
End Property

' -----------------------------------------------------------------------------
' ListenKeyboard (PropertyGet), source lines 641-643
' -----------------------------------------------------------------------------
Private Property Get ListenKeyboard() As Boolean
    ListenKeyboard = lListenKeyboard
End Property

' -----------------------------------------------------------------------------
' Deployed (PropertyLet), source lines 644-651
' -----------------------------------------------------------------------------
Private Property Let Deployed(oValue As Boolean)
    lDeployed = oValue
    If lDeployed Then
        PlotterForm.Height = 657.75
    Else
        PlotterForm.Height = 525
    End If
End Property

' -----------------------------------------------------------------------------
' Deployed (PropertyGet), source lines 652-654
' -----------------------------------------------------------------------------
Private Property Get Deployed() As Boolean
    Deployed = lDeployed
End Property

' -----------------------------------------------------------------------------
' Combo1_Click (Sub), source lines 657-659
' -----------------------------------------------------------------------------
Private Sub Combo1_Click()
    ListenKeyboard = False
End Sub

' -----------------------------------------------------------------------------
' ComboBox2_Click (Sub), source lines 661-663
' -----------------------------------------------------------------------------
Private Sub ComboBox2_Click()
    ListenKeyboard = False
End Sub

' -----------------------------------------------------------------------------
' ComboBox3_Click (Sub), source lines 665-667
' -----------------------------------------------------------------------------
Private Sub ComboBox3_Click()
    ListenKeyboard = False
End Sub

' -----------------------------------------------------------------------------
' CommandButton11_MouseDown (Sub), source lines 671-682
' -----------------------------------------------------------------------------
Private Sub CommandButton11_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim l As Double:    l = 5
    MovingProcesing = True
    DoEvents
    Do While MovingProcesing
        OutPuts.MoveUP l
        RefreshXY
        Sleep 140
        If l < 10 Then l = l + 5
        DoEvents
    Loop
End Sub

' -----------------------------------------------------------------------------
' CommandButton12_MouseDown (Sub), source lines 683-694
' -----------------------------------------------------------------------------
Private Sub CommandButton12_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim l As Double:    l = 5
    MovingProcesing = True
    DoEvents
    Do While MovingProcesing
        OutPuts.MoveDown l
        RefreshXY
        Sleep 140
        If l < 10 Then l = l + 5
        DoEvents
    Loop
End Sub

' -----------------------------------------------------------------------------
' CommandButton13_MouseDown (Sub), source lines 695-706
' -----------------------------------------------------------------------------
Private Sub CommandButton13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim l As Double:    l = 5
    MovingProcesing = True
    DoEvents
    Do While MovingProcesing
        OutPuts.MoveLeft l
        RefreshXY
        Sleep 140
        If l < 10 Then l = l + 5
        DoEvents
    Loop
End Sub

' -----------------------------------------------------------------------------
' CommandButton14_MouseDown (Sub), source lines 707-718
' -----------------------------------------------------------------------------
Private Sub CommandButton14_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim l As Double:    l = 5
    MovingProcesing = True
    DoEvents
    Do While MovingProcesing
        OutPuts.MoveRight l
        RefreshXY
        Sleep 140
        If l < 10 Then l = l + 5
        DoEvents
    Loop
End Sub

' -----------------------------------------------------------------------------
' CommandButton13_MouseUp (Sub), source lines 719-721
' -----------------------------------------------------------------------------
Private Sub CommandButton13_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MovingProcesing = False
End Sub

' -----------------------------------------------------------------------------
' CommandButton14_MouseUp (Sub), source lines 722-724
' -----------------------------------------------------------------------------
Private Sub CommandButton14_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MovingProcesing = False
End Sub

' -----------------------------------------------------------------------------
' CommandButton11_MouseUp (Sub), source lines 725-727
' -----------------------------------------------------------------------------
Private Sub CommandButton11_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MovingProcesing = False
End Sub

' -----------------------------------------------------------------------------
' CommandButton12_MouseUp (Sub), source lines 728-730
' -----------------------------------------------------------------------------
Private Sub CommandButton12_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MovingProcesing = False
End Sub

' -----------------------------------------------------------------------------
' RefreshXY (Sub), source lines 731-747
' -----------------------------------------------------------------------------
Private Sub RefreshXY()
    Label18.Caption = "X = " & CStr(Round(OutPuts.Xmm, 0)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    Label19.Caption = "Y = " & CStr(Round(OutPuts.Ymm, 0)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    
    If Not PltDocument1(TabStrip1.value) Is Nothing Then
        Dim p As POINTL
        p.x = OutPuts.Xmm * 1000
        p.y = OutPuts.Ymm * 1000
        PltDocument1(TabStrip1.value).XYmargin = p
    End If

'    If OutPuts.x = 0 And OutPuts.y = 0 Then
'        CommandButton21.Enabled = False
'    Else
'        CommandButton21.Enabled = True
'    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton21_Click (Sub), source lines 750-753
' -----------------------------------------------------------------------------
Private Sub CommandButton21_Click()
    OutPuts.Move00
    RefreshXY
End Sub

' -----------------------------------------------------------------------------
' CommandButton22_Click (Sub), source lines 756-760
' -----------------------------------------------------------------------------
Private Sub CommandButton22_Click()
'    OutPuts.Out PlotterLanguage.StrToAscii("IN;PA;VS5;PU0,0;PU0,400;PU400,400;PU400,0;PU0,0;"), 0, 0
    OutPuts.test
    RefreshXY
End Sub

' -----------------------------------------------------------------------------
' CommandButton23_Click (Sub), source lines 766-768
' -----------------------------------------------------------------------------
Private Sub CommandButton23_Click()
    ListenKeyboard = Not ListenKeyboard
End Sub

' -----------------------------------------------------------------------------
' TextBox10_Enter (Sub), source lines 772-774
' -----------------------------------------------------------------------------
Private Sub TextBox10_Enter()
    ListenKeyboard = False
End Sub

' -----------------------------------------------------------------------------
' TextBox12_Enter (Sub), source lines 776-778
' -----------------------------------------------------------------------------
Private Sub TextBox12_Enter()
'    ListenKeyboard = False
End Sub

' -----------------------------------------------------------------------------
' TextBox9_Enter (Sub), source lines 780-782
' -----------------------------------------------------------------------------
Private Sub TextBox9_Enter()
    ListenKeyboard = False
End Sub

' -----------------------------------------------------------------------------
' KeyReadLoop (Sub), source lines 785-868
' -----------------------------------------------------------------------------
Private Sub KeyReadLoop()
    On Error Resume Next
    KeyLoopProcessing = True
    Dim l As Double:  l = 5
    Dim pKey As PressedKeys
    Static lastTimer As Single
    Static lastTimer2 As Single
    
    Do While KeyLoopProcessing
        pKey = PressedKey
        If ListenKeyboard Then
            If Timer - lastTimer > 0.1 Then
                If CommandButton23.Picture = 0 Then
                    CommandButton23.Picture = Image3.Picture
                Else
                    CommandButton23.Picture = Nothing
                End If
                lastTimer = Timer
            End If
            Select Case pKey
                Case KeyLeft
                    OutPuts.MoveLeft l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyRight
                    OutPuts.MoveRight l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyUp
                    OutPuts.MoveUP l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyDown
                    OutPuts.MoveDown l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyLeftUp
                    OutPuts.MoveLeftUp l
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyUpRight
                    OutPuts.MoveUpRight
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyRightDown
                    OutPuts.MoveRightDown
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyDownLeft
                    OutPuts.MoveDownLeft
                    RefreshXY
                    If l < 10 Then l = l + 5
                Case KeyEnter
                    
                Case Key0
                    OutPuts.Move00
                    RefreshXY
                Case Else
                    l = 5
            End Select
        End If

        If Not USB_Events1 Is Nothing Then
            USB_Events1.USB_change_check_loop
            'Chek wifi connection
            OutPuts.WiFi_change_check_loop
        End If
        
        If Label22.Visible Then
            If Timer - lastTimer2 > 0.1 Then
                Label23.Visible = Not Label23.Visible
                lastTimer2 = Timer
            End If
        Else
            Label23.Visible = False
        End If


        For q = 0 To 6
            Sleep 20
            DoEvents
        Next
    Loop
End Sub

' -----------------------------------------------------------------------------
' PressedKey (Function), source lines 869-941
' -----------------------------------------------------------------------------
Private Function PressedKey() As PressedKeys

    Static myKey As Integer
    Dim mL As Boolean
    Dim mR As Boolean
    Dim mU As Boolean
    Dim mD As Boolean
    Static mLeft As Boolean
    Static mRight As Boolean
    Static mUp As Boolean
    Static mDown As Boolean
    
    
    myKey = GetAsyncKeystate(vbKeyUp)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mU = True
    
    myKey = GetAsyncKeystate(vbKeyLeft)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mL = True

    myKey = GetAsyncKeystate(vbKeyRight)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mR = True

    myKey = GetAsyncKeystate(vbKeyDown)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mD = True

    myKey = GetAsyncKeystate(vbKeyNumpad4)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mL = True

    myKey = GetAsyncKeystate(vbKeyNumpad6)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mR = True

    myKey = GetAsyncKeystate(vbKeyNumpad8)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mU = True

    myKey = GetAsyncKeystate(vbKeyNumpad2)
    If (myKey And -32768) = -32768 Or (myKey And &H1) = &H1 Then mD = True

    mLeft = mL And (Not mR)
    mRight = mR And (Not mL)
    mUp = mU And (Not mD)
    mDown = mD And (Not mU)
    
    
    If mLeft And mUp Then PressedKey = KeyLeftUp: Exit Function
    If mUp And mRight Then PressedKey = KeyUpRight: Exit Function
    If mRight And mDown Then PressedKey = KeyRightDown: Exit Function
    If mDown And mLeft Then PressedKey = KeyDownLeft: Exit Function
    
    If mUp Then PressedKey = KeyUp: Exit Function
    If mLeft Then PressedKey = KeyLeft: Exit Function
    If mRight Then PressedKey = KeyRight: Exit Function
    If mDown Then PressedKey = KeyDown: Exit Function
    
    
    If (GetAsyncKeystate(vbKey0) And &H1) = &H1 Then
        PressedKey = Key0
        Exit Function
    End If
    If (GetAsyncKeystate(vbKeyNumpad0) And &H1) = &H1 Then
        PressedKey = Key0
        Exit Function
    End If
    If (GetAsyncKeystate(vbKeyInsert) And &H1) = &H1 Then
        PressedKey = Key0
        Exit Function
    End If
    
    If (GetAsyncKeystate(vbKeyReturn) And &H1) = &H1 Then
        PressedKey = KeyEnter
        Exit Function
    End If

End Function

' -----------------------------------------------------------------------------
' TabStrip1_Change (Sub), source lines 945-969
' -----------------------------------------------------------------------------
Private Sub TabStrip1_Change()
    RegData1.ActiveProfile = TabStrip1.value
    StopShowAll
    If TabStrip1.SelectedItem.index = 8 Then
        Frame1.Visible = True
        Frame2.Visible = False
        
        TextBox11.Enabled = False
        Label20.Enabled = False
        TextBox12.Enabled = False
        CommandButton20.Enabled = False
    Else
        Frame1.Visible = False
        Frame2.Visible = True
        If Initialized Then
            RefreshFields
        End If
        
        TextBox11.Enabled = True
        Label20.Enabled = True
        TextBox12.Enabled = True
        TextBox12_Change
        RefreshPltDocument
    End If
End Sub

' -----------------------------------------------------------------------------
' RefreshFields (Sub), source lines 970-1075
' -----------------------------------------------------------------------------
Public Sub RefreshFields()
    If TabStrip1.SelectedItem.index = 8 Then Exit Sub
    
    With Profiles(TabStrip1.SelectedItem.index)
        Text4.text = .Device
        Combo1.ListIndex = ListIndexByName(.OutPut, Combo1)
        TextBox9.text = Format((.ClosingCurve / 1000), "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        TextBox10.text = Format((.KnifeOffset / 1000), "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        CheckBox9.value = RegData1.PreRun
        CheckBox10.value = .DoublePass
        Dim q As Long
        '====Accelerations====
        With ComboBox1
            .Clear
            .AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            For q = 0 To Profiles(TabStrip1.SelectedItem.index).AccelerationsCount - 1
                .AddItem Replace(Profiles(TabStrip1.SelectedItem.index).Accelerations(q) / 10 & " " & localization.ActiveLng.Item("PlotterForm.m/s") & ChrW(178), ",", ".")
            Next
            
            If .ListCount = 1 Or (Profiles(TabStrip1.SelectedItem.index).CommandLanguage = DMPL) Then
                .Enabled = False
                .ListIndex = 0
                Label14.Enabled = False
                ImageAcceleration.Picture = ImageAccelerationDisabled.Picture
            Else
                .ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveAcceleration
                .Enabled = True
                Label14.Enabled = True
                ImageAcceleration.Picture = ImageAccelerationEnabled.Picture
            End If
        End With
        '====Speeds====
        With ComboBox2
            If Profiles(TabStrip1.SelectedItem.index).SeparateParametrs Then
                .Visible = False
                ComboBox6.Visible = True
                ComboBox7.Visible = True
            Else
                .Visible = True
                ComboBox6.Visible = False
                ComboBox7.Visible = False
            End If
            
            .Clear
            ComboBox6.Clear
            ComboBox7.Clear
            
            .AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            ComboBox6.AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            ComboBox7.AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            
            For q = 0 To Profiles(TabStrip1.SelectedItem.index).SpeedsCount - 1
                .AddItem Profiles(TabStrip1.SelectedItem.index).Speeds(q) & " " & localization.ActiveLng.Item("PlotterForm.mm/s")
                ComboBox6.AddItem ChrW(9660) & " " & Profiles(TabStrip1.SelectedItem.index).Speeds(q) & " " & localization.ActiveLng.Item("PlotterForm.mm/s")
                ComboBox7.AddItem ChrW(9650) & " " & Profiles(TabStrip1.SelectedItem.index).Speeds(q) & " " & localization.ActiveLng.Item("PlotterForm.mm/s")
            Next
            If .ListCount = 1 Then
                .ListIndex = 0
                ComboBox6.ListIndex = 0
                ComboBox7.ListIndex = 0
                .Enabled = False
                ComboBox6.Enabled = False
                ComboBox7.Enabled = False
                Label12.Enabled = False
                ImageSpeed.Picture = ImageSpeedDisebled.Picture
            Else
                .ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveSpeed
                ComboBox6.ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveSpeed
                ComboBox7.ListIndex = Profiles(TabStrip1.SelectedItem.index).ActiveSpeedUP
                
                .Enabled = True
                ComboBox6.Enabled = True
                ComboBox7.Enabled = True
                Label12.Enabled = True

                ImageSpeed.Picture = ImageSpeedEnabled.Picture
            End If
        End With
        
        '====Forces====
        With ComboBox3
            .Clear
            .AddItem localization.ActiveLng.Item("PlotterForm.ComboBox1")
            For q = 0 To Profiles(TabStrip1.SelectedItem.index).ForcesCount - 1
                .AddItem Profiles(TabStrip1.SelectedItem.index).Forces(q) & " " & localization.ActiveLng.Item("PlotterForm.g")
            Next
            If .ListCount = 1 Then
                .Enabled = False
                .ListIndex = 0
                Label13.Enabled = False
                ImageForce.Picture = ImageForceDisabled.Picture
            Else
                .ListIndex = ListIndexByName(Profiles(TabStrip1.SelectedItem.index).ActiveForce & " " & localization.ActiveLng.Item("PlotterForm.g"), ComboBox3)
                .Enabled = True
                Label13.Enabled = True
                ImageForce.Picture = ImageForceEnabled.Picture
            End If
        End With
        If ComboBox1.Enabled = False And ComboBox2.Enabled = False And ComboBox3.Enabled = False Then
            Frame4.Enabled = False
        Else
            Frame4.Enabled = True
        End If
    End With
    RefreshXY
End Sub

' -----------------------------------------------------------------------------
' Combo1_Change (Sub), source lines 1076-1140
' -----------------------------------------------------------------------------
Private Sub Combo1_Change()
    If TabStrip1.SelectedItem.index > 7 Then Exit Sub
    If suppressChange Then Exit Sub
    Profiles(TabStrip1.SelectedItem.index).OutPut = KillBirdsAndSpaces(Combo1.text)
    RefreshXY
    If Combo1.ListIndex + 1 = Combo1.ListCount Then
        CommandButton22.Enabled = False
        If Not OriginalCurves Is Nothing Then
            If OriginalCurves.count = 0 Then
                Command5.Enabled = False
            Else
                Command5.Enabled = True
            End If
        End If
        Label22.Visible = False

        
        CommandButton21.Enabled = True
        CommandButton11.Enabled = True
        CommandButton13.Enabled = True
        CommandButton12.Enabled = True
        CommandButton14.Enabled = True
        
        Combo1.BackColor = &H80000005
    Else
        If OutPuts.Connected(Combo1.ListIndex) = True Then
            CommandButton22.Enabled = True

            If Not OriginalCurves Is Nothing Then
                If OriginalCurves.count > 0 Then
                    Command5.Enabled = True
                Else
                    Command5.Enabled = False
                End If
            End If

            Label22.Visible = False
            Label23.Visible = False
            
            CommandButton21.Enabled = True
            CommandButton11.Enabled = True
            CommandButton13.Enabled = True
            CommandButton12.Enabled = True
            CommandButton14.Enabled = True
            
            Combo1.BackColor = &H80000005
        Else
            CommandButton22.Enabled = False
            Command5.Enabled = False
            Label22.Visible = True
            
            
            CommandButton21.Enabled = False
            CommandButton11.Enabled = False
            CommandButton13.Enabled = False
            CommandButton12.Enabled = False
            CommandButton14.Enabled = False
            
            Combo1.BackColor = &HFFFF&
        End If
    End If
    Combo1.ControlTipText = KillBirdsAndSpaces(Combo1.text)
    TextBox12_Change
    RefreshStartButton
End Sub

' -----------------------------------------------------------------------------
' ComboBox1_Change (Sub), source lines 1141-1147
' -----------------------------------------------------------------------------
Private Sub ComboBox1_Change()
    With ComboBox1
        If .ListIndex > -1 Then
            Profiles(TabStrip1.SelectedItem.index).ActiveAcceleration = .ListIndex
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' ComboBox6_Change (Sub), source lines 1148-1159
' -----------------------------------------------------------------------------
Private Sub ComboBox6_Change()
    With ComboBox6
        If .ListIndex > -1 Then
            If .ListIndex = 0 Then
                ComboBox7.ListIndex = 0
            Else
                If ComboBox7.ListIndex = 0 Then ComboBox7.ListIndex = .ListIndex
            End If
            Profiles(TabStrip1.SelectedItem.index).ActiveSpeed = .ListIndex
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' ComboBox7_Change (Sub), source lines 1160-1171
' -----------------------------------------------------------------------------
Private Sub ComboBox7_Change()
    With ComboBox7
        If .ListIndex > -1 Then
            If .ListIndex = 0 Then
                ComboBox6.ListIndex = 0
            Else
                If ComboBox6.ListIndex = 0 Then ComboBox6.ListIndex = .ListIndex
            End If
            Profiles(TabStrip1.SelectedItem.index).ActiveSpeedUP = .ListIndex
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' ComboBox2_Change (Sub), source lines 1172-1179
' -----------------------------------------------------------------------------
Private Sub ComboBox2_Change()
    With ComboBox2
        If .ListIndex > -1 Then
            Profiles(TabStrip1.SelectedItem.index).ActiveSpeed = .ListIndex
            ComboBox6.ListIndex = .ListIndex
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' ComboBox3_Change (Sub), source lines 1180-1190
' -----------------------------------------------------------------------------
Private Sub ComboBox3_Change()
    With ComboBox3
        If .ListIndex > -1 Then
            If .ListIndex = 0 Then
                Profiles(TabStrip1.SelectedItem.index).ActiveForce = 0
            Else
                Profiles(TabStrip1.SelectedItem.index).ActiveForce = Profiles(TabStrip1.SelectedItem.index).Forces(.ListIndex - 1)
            End If
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' CheckBox1_Change (Sub), source lines 1191-1197
' -----------------------------------------------------------------------------
Private Sub CheckBox1_Change()
    Dim np As Boolean: np = Not CheckBox1.value
    TabStrip1.Tabs(0).Visible = np
    TextBox1.Enabled = np
    Label1.Enabled = np
    Profiles(0).Visible = np
End Sub

' -----------------------------------------------------------------------------
' CheckBox2_Change (Sub), source lines 1198-1204
' -----------------------------------------------------------------------------
Private Sub CheckBox2_Change()
    Dim np As Boolean: np = Not CheckBox2.value
    TabStrip1.Tabs(1).Visible = np
    TextBox2.Enabled = np
    Label2.Enabled = np
    Profiles(1).Visible = np
End Sub

' -----------------------------------------------------------------------------
' CheckBox3_Change (Sub), source lines 1205-1211
' -----------------------------------------------------------------------------
Private Sub CheckBox3_Change()
    Dim np As Boolean: np = Not CheckBox3.value
    TabStrip1.Tabs(2).Visible = np
    TextBox3.Enabled = np
    Label3.Enabled = np
    Profiles(2).Visible = np
End Sub

' -----------------------------------------------------------------------------
' CheckBox4_Change (Sub), source lines 1212-1218
' -----------------------------------------------------------------------------
Private Sub CheckBox4_Change()
    Dim np As Boolean: np = Not CheckBox4.value
    TabStrip1.Tabs(3).Visible = np
    TextBox4.Enabled = np
    Label4.Enabled = np
    Profiles(3).Visible = np
End Sub

' -----------------------------------------------------------------------------
' CheckBox5_Change (Sub), source lines 1219-1225
' -----------------------------------------------------------------------------
Private Sub CheckBox5_Change()
    Dim np As Boolean: np = Not CheckBox5.value
    TabStrip1.Tabs(4).Visible = np
    TextBox5.Enabled = np
    Label5.Enabled = np
    Profiles(4).Visible = np
End Sub

' -----------------------------------------------------------------------------
' CheckBox6_Change (Sub), source lines 1226-1232
' -----------------------------------------------------------------------------
Private Sub CheckBox6_Change()
    Dim np As Boolean: np = Not CheckBox6.value
    TabStrip1.Tabs(5).Visible = np
    TextBox6.Enabled = np
    Label6.Enabled = np
    Profiles(5).Visible = np
End Sub

' -----------------------------------------------------------------------------
' CheckBox7_Change (Sub), source lines 1233-1239
' -----------------------------------------------------------------------------
Private Sub CheckBox7_Change()
    Dim np As Boolean: np = Not CheckBox7.value
    TabStrip1.Tabs(6).Visible = np
    TextBox7.Enabled = np
    Label7.Enabled = np
    Profiles(6).Visible = np
End Sub

' -----------------------------------------------------------------------------
' CheckBox8_Change (Sub), source lines 1240-1246
' -----------------------------------------------------------------------------
Private Sub CheckBox8_Change()
    Dim np As Boolean: np = Not CheckBox8.value
    TabStrip1.Tabs(7).Visible = np
    TextBox8.Enabled = np
    Label8.Enabled = np
    Profiles(7).Visible = np
End Sub

' -----------------------------------------------------------------------------
' TextBox1_Change (Sub), source lines 1247-1250
' -----------------------------------------------------------------------------
Private Sub TextBox1_Change()
    TabStrip1.Tabs(0).Caption = TextBox1.text
    Profiles(0).name = TextBox1.text
End Sub

' -----------------------------------------------------------------------------
' TextBox2_Change (Sub), source lines 1251-1254
' -----------------------------------------------------------------------------
Private Sub TextBox2_Change()
    TabStrip1.Tabs(1).Caption = TextBox2.text
    Profiles(1).name = TextBox2.text
End Sub

' -----------------------------------------------------------------------------
' TextBox3_Change (Sub), source lines 1255-1258
' -----------------------------------------------------------------------------
Private Sub TextBox3_Change()
    TabStrip1.Tabs(2).Caption = TextBox3.text
    Profiles(2).name = TextBox3.text
End Sub

' -----------------------------------------------------------------------------
' TextBox4_Change (Sub), source lines 1259-1262
' -----------------------------------------------------------------------------
Private Sub TextBox4_Change()
    TabStrip1.Tabs(3).Caption = TextBox4.text
    Profiles(3).name = TextBox4.text
End Sub

' -----------------------------------------------------------------------------
' TextBox5_Change (Sub), source lines 1263-1266
' -----------------------------------------------------------------------------
Private Sub TextBox5_Change()
    TabStrip1.Tabs(4).Caption = TextBox5.text
    Profiles(4).name = TextBox5.text
End Sub

' -----------------------------------------------------------------------------
' TextBox6_Change (Sub), source lines 1267-1270
' -----------------------------------------------------------------------------
Private Sub TextBox6_Change()
    TabStrip1.Tabs(5).Caption = TextBox6.text
    Profiles(5).name = TextBox6.text
End Sub

' -----------------------------------------------------------------------------
' TextBox7_Change (Sub), source lines 1271-1274
' -----------------------------------------------------------------------------
Private Sub TextBox7_Change()
    TabStrip1.Tabs(6).Caption = TextBox7.text
    Profiles(6).name = TextBox7.text
End Sub

' -----------------------------------------------------------------------------
' TextBox8_Change (Sub), source lines 1275-1278
' -----------------------------------------------------------------------------
Private Sub TextBox8_Change()
    TabStrip1.Tabs(7).Caption = TextBox8.text
    Profiles(7).name = TextBox8.text
End Sub

' -----------------------------------------------------------------------------
' CheckBox9_Change (Sub), source lines 1279-1282
' -----------------------------------------------------------------------------
Private Sub CheckBox9_Change()
    RegData1.PreRun = CheckBox9.value
    RefreshStartButton
End Sub

' -----------------------------------------------------------------------------
' RefreshStartButton (Sub), source lines 1284-1298
' -----------------------------------------------------------------------------
Private Sub RefreshStartButton()
    If Combo1.ListIndex + 1 = Combo1.ListCount Then
        Command5.Caption = localization.ActiveLng.Item("PlotterForm.Command5.Caption_SaveAs")
        CheckBox9.Enabled = False
    Else
        If CheckBox9.value Then
            Command5.Caption = localization.ActiveLng.Item("PlotterForm.Command5.Caption_Pre-run")
        Else
            Command5.Caption = localization.ActiveLng.Item("PlotterForm.Command5.Caption_Run")
        End If

        CheckBox9.Enabled = True

    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton1_Click (Sub), source lines 1300-1350
' -----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    
    ListenKeyboard = False
    PltDocument1(TabStrip1.value).StopShowWay
    DevicesForm.Show 1
    'Set Changes
    With PltDocument1(TabStrip1.value)
        If Profiles(TabStrip1.value).MaxSheetHeight = 0 Then
            .MaxSheetHeight = 2147483647
        Else
            .MaxSheetHeight = Profiles(TabStrip1.value).MaxSheetHeight * 1000
        End If
        
        If Profiles(TabStrip1.value).MaxSheetWidth = 0 Then
            If Profiles(TabStrip1.value).Turn = Turn0 Or Profiles(TabStrip1.value).Turn = Turn180 Then
                If OriginalWidth <= 600000 Then
                    .MaxSheetWidth = 600000
                ElseIf OriginalWidth <= 1600000 Then
                    .MaxSheetWidth = 1600000
                Else
                    .MaxSheetWidth = OriginalWidth
                End If
            Else
                If OriginalHeight <= 600000 Then
                    .MaxSheetWidth = 600000
                ElseIf OriginalHeight <= 1600000 Then
                    .MaxSheetWidth = 1600000
                Else
                    .MaxSheetWidth = OriginalHeight
                End If
            End If
        Else
            .MaxSheetWidth = Profiles(TabStrip1.value).MaxSheetWidth * 1000
        End If

'            Set .MaterialBrush = CreateBrush(RGB(MaxColor.R, MaxColor.G, MaxColor.B), CSng(MaxColor.A) / 255 * 100)

        .PlotterModelString = Profiles(TabStrip1.value).PlotterModelString
        .PlotterName = Profiles(TabStrip1.value).Device

        .SetCurvesByVal OriginalCurves, OriginalHeight, OriginalWidth, Profiles(TabStrip1.value).Turn, Profiles(TabStrip1.value).HeightMirror, Profiles(TabStrip1.value).WidthMirror
        .ZeroPosition = Profiles(TabStrip1.value).ZeroPosition
        
        Dim p As POINTL
        p.x = OutPuts.Xmm * 1000
        p.y = OutPuts.Ymm * 1000
        PltDocument1(TabStrip1.value).XYmargin = p
        
        .Render
    End With
End Sub

' -----------------------------------------------------------------------------
' CommandButton2_Click (Sub), source lines 1351-1354
' -----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    ListenKeyboard = False
    OutPuts.Configure KillBirdsAndSpaces(Combo1.text), FindWindow(vbNullString, Me.Caption)
End Sub

' -----------------------------------------------------------------------------
' TextBox9_Change (Sub), source lines 1355-1359
' -----------------------------------------------------------------------------
Private Sub TextBox9_Change()
    On Error GoTo 10
    Profiles(TabStrip1.SelectedItem.index).ClosingCurve = ValP(TextBox9.text) * 1000
10
End Sub

' -----------------------------------------------------------------------------
' TextBox9_AfterUpdate (Sub), source lines 1360-1368
' -----------------------------------------------------------------------------
Private Sub TextBox9_AfterUpdate()
    On Error GoTo 10
    Dim d As Long
    d = ((ValP(TextBox9.text) * 1000 + 12.5) \ 25) * 25
    If d < 0 Then d = 0
    If d > 10000 Then d = 10000
10
    TextBox9.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub

' -----------------------------------------------------------------------------
' TextBox10_Change (Sub), source lines 1369-1373
' -----------------------------------------------------------------------------
Private Sub TextBox10_Change()
    On Error GoTo 10
    Profiles(TabStrip1.SelectedItem.index).KnifeOffset = ValP(TextBox10.text) * 1000
10
End Sub

' -----------------------------------------------------------------------------
' TextBox10_AfterUpdate (Sub), source lines 1374-1382
' -----------------------------------------------------------------------------
Private Sub TextBox10_AfterUpdate()
    On Error GoTo 10
    Dim d As Long
    d = ((ValP(TextBox10.text) * 1000 + 12.5) \ 25) * 25
    If d < 0 Then d = 0
    If d > 5550 Then d = 5550
10
    TextBox10.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub

' -----------------------------------------------------------------------------
' CommandButton3_Click (Sub), source lines 1383-1387
' -----------------------------------------------------------------------------
Private Sub CommandButton3_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(0).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton4_Click (Sub), source lines 1389-1393
' -----------------------------------------------------------------------------
Private Sub CommandButton4_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(1).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton5_Click (Sub), source lines 1395-1399
' -----------------------------------------------------------------------------
Private Sub CommandButton5_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(2).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton6_Click (Sub), source lines 1401-1405
' -----------------------------------------------------------------------------
Private Sub CommandButton6_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(3).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton7_Click (Sub), source lines 1407-1411
' -----------------------------------------------------------------------------
Private Sub CommandButton7_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(4).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton8_Click (Sub), source lines 1413-1417
' -----------------------------------------------------------------------------
Private Sub CommandButton8_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(5).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton9_Click (Sub), source lines 1419-1423
' -----------------------------------------------------------------------------
Private Sub CommandButton9_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(6).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton10_Click (Sub), source lines 1424-1428
' -----------------------------------------------------------------------------
Private Sub CommandButton10_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & Profiles(7).file, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton19_Click (Sub), source lines 1430-1440
' -----------------------------------------------------------------------------
Private Sub CommandButton19_Click()
    Dim d As Long
    d = ((ValP(TextBox9.text) * 1000) \ 25) * 25
    d = d + 25
    If d < 0 Then d = 0
    If d > 10000 Then d = 10000
    TextBox9.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox9.SetFocus
    TextBox9.SelStart = 0
    TextBox9.SelLength = Len(TextBox9.text) - 3
End Sub

' -----------------------------------------------------------------------------
' CommandButton18_Click (Sub), source lines 1441-1451
' -----------------------------------------------------------------------------
Private Sub CommandButton18_Click()
    Dim d As Long
    d = ((ValP(TextBox9.text) * 1000 + 24) \ 25) * 25
    d = d - 25
    If d < 0 Then d = 0
    If d > 10000 Then d = 10000
    TextBox9.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox9.SetFocus
    TextBox9.SelStart = 0
    TextBox9.SelLength = 5
End Sub

' -----------------------------------------------------------------------------
' CommandButton16_Click (Sub), source lines 1452-1462
' -----------------------------------------------------------------------------
Private Sub CommandButton16_Click()
    Dim d As Long
    d = ((ValP(TextBox10.text) * 1000) \ 25) * 25
    d = d + 25
    If d < 0 Then d = 0
    If d > 5550 Then d = 5550
    TextBox10.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox10.SetFocus
    TextBox10.SelStart = 2
    TextBox10.SelLength = 3
End Sub

' -----------------------------------------------------------------------------
' CommandButton17_Click (Sub), source lines 1463-1473
' -----------------------------------------------------------------------------
Private Sub CommandButton17_Click()
    Dim d As Long
    d = ((ValP(TextBox10.text) * 1000 + 24) \ 25) * 25
    d = d - 25
    If d < 0 Then d = 0
    If d > 5550 Then d = 5550
    TextBox10.text = Format(d / 1000, "0.000") & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    TextBox10.SetFocus
    TextBox10.SelStart = 2
    TextBox10.SelLength = 3
End Sub

' -----------------------------------------------------------------------------
' ValP (Function), source lines 1474-1476
' -----------------------------------------------------------------------------
Private Function ValP(ByRef txt As String) As Double
    ValP = Abs(val(Replace$(txt, ",", ".")))
End Function

' -----------------------------------------------------------------------------
' CommandButton15_Click (Sub), source lines 1477-1480
' -----------------------------------------------------------------------------
Private Sub CommandButton15_Click()
    Deployed = Not Deployed
    RegData1.Deployed = Deployed
End Sub

' -----------------------------------------------------------------------------
' AddLog (Sub), source lines 1481-1495
' -----------------------------------------------------------------------------
Public Sub AddLog(txt As String, Optional isInput As Boolean = False)
    On Error GoTo 10
    TextBox11.SetFocus
    txt = Replace(txt, vbNewLine, "")
    txt = Trim(txt)
    If txt = "" Then GoTo 10
    
    If isInput Then
        TextBox11.text = TextBox11.text & KillBirdsAndSpaces(Combo1.text) & " >> "
    Else
        TextBox11.text = TextBox11.text & KillBirdsAndSpaces(Combo1.text) & " << "
    End If
    TextBox11.text = TextBox11.text & txt & vbNewLine
10
End Sub

' -----------------------------------------------------------------------------
' CommandButton20_Click (Sub), source lines 1496-1506
' -----------------------------------------------------------------------------
Private Sub CommandButton20_Click()
    Dim txt As String
    txt = UCase(TextBox12.text)
    Dim bytes() As Byte
    bytes = StrToAscii(txt)
    OutPuts.Out bytes, OutPuts.x, OutPuts.y
    SendKeys "{PGDN}"
    TextBox12.text = ""
    TextBox12.SetFocus
    
End Sub

' -----------------------------------------------------------------------------
' TextBox12_Change (Sub), source lines 1507-1518
' -----------------------------------------------------------------------------
Private Sub TextBox12_Change()
    Dim txt As String
    txt = Replace(TextBox12.text, vbNewLine, "")
    txt = Trim(txt)
    If Combo1.ListIndex > -1 Then
        If Len(txt) > 0 And OutPuts.Connected(Combo1.ListIndex) = True Then
            CommandButton20.Enabled = True
        Else
            CommandButton20.Enabled = False
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' TextBox12_KeyUp (Sub), source lines 1519-1521
' -----------------------------------------------------------------------------
Private Sub TextBox12_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then CommandButton20_Click
End Sub

' -----------------------------------------------------------------------------
' UserForm_Layout (Sub), source lines 1526-1531
' -----------------------------------------------------------------------------
Private Sub UserForm_Layout()
    If TabStrip1.value > 7 Then Exit Sub
    If Not PltDocument1(TabStrip1.value) Is Nothing Then
        PltDocument1(TabStrip1.value).Refresh
    End If
End Sub

' -----------------------------------------------------------------------------
' KillBirdsAndSpaces (Function), source lines 1532-1551
' -----------------------------------------------------------------------------
Private Function KillBirdsAndSpaces(ByVal s As String) As String
    ' Удалить 4 пробела в начале, если строка достаточно длинная
    If Len(s) >= 4 Then
        If left(s, 4) = "    " Then
            KillBirdsAndSpaces = Mid(s, 5)
            Exit Function
        End If
    End If

    ' Удалить символ ChrW(10003) (?) в начале, если строка не пустая
    If Len(s) >= 1 Then
        If left(s, 1) = ChrW(10003) Then
            KillBirdsAndSpaces = Mid(s, 2)
            Exit Function
        End If
    End If

    ' Вернуть как есть, если ничего не найдено
    KillBirdsAndSpaces = s
End Function

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 1552-1563
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    KeyLoopProcessing = False
    If Not USB_Events1 Is Nothing Then USB_Events1.Terminate
    If Not BlocksSocket1 Is Nothing Then BlocksSocket1.Disconnect
    Set OutPuts = Nothing
    Dim q As Long
    For q = 0 To CountProfiles - 1
        Set Profiles(q) = Nothing
    Next
    Converter.Break
    StopShowAll
End Sub
