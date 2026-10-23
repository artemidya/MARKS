' =============================================================================
' Group: 02_Calculation_and_Nesting_Forms
' Source module: RaskroyForm.frm
' Procedures: 104
' =============================================================================

' -----------------------------------------------------------------------------
' CheckBox1_Change (Sub), source lines 233-237
' -----------------------------------------------------------------------------
Private Sub CheckBox1_Change()
    If Not dntChangeCheckBox1 Then
        RegData1.NestingPlaceOnPage = CheckBox1.value
    End If
End Sub

' -----------------------------------------------------------------------------
' Combo1_KeyDown (Sub), source lines 240-242
' -----------------------------------------------------------------------------
Private Sub Combo1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' Command1_KeyDown (Sub), source lines 243-245
' -----------------------------------------------------------------------------
Private Sub Command1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 250-305
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    UserData.Load
    Dim kdpi As Double
    kdpi = MonitorDPI.GetDpiDbl
    
    shiftX = 360 * kdpi
    shiftY = 15 * kdpi
    WindowWidth = 806 * kdpi
    WindowHeigth = 592 * kdpi

    lastUnit = Application.ActiveDocument.Unit
    
    lastReferencePoint = ActiveDocument.ReferencePoint
    ActiveDocument.ReferencePoint = cdrTopLeft
    Application.ActiveDocument.Unit = cdrInch

    Combo2.AddItem ' "без поворота", 0
    Combo2.AddItem "180°", 1
    Combo2.AddItem "90°", 2
    Combo2.AddItem "45°", 3
    Combo2.AddItem "5°", 4
    SetTexts
    ProtocolStatus = NotConnected


    'sorting
    UserData.SortProfiles
    
    Combo2.ListIndex = 0
    
    Dim q As Long
    For q = 0 To UserData.CountRaskroyProfiles - 1
        Combo1.AddItem Trim(UserData.RaskroyProfiles(q).name), q
    Next
    
    NeedSave = False
    Plotf = True
    
    If UserData.LastProfileSelectionNumber <= (UserData.CountRaskroyProfiles - 1) Then
        Combo1.ListIndex = UserData.LastProfileSelectionNumber
    End If
        
    If UserData.CountRaskroyProfiles = 0 Then
        Combo1.text = " " & localization.ActiveLng.Item("RaskroyForm.empty")
        Combo1.Enabled = False
        CommandDelete.Enabled = False
        CommandSave.Enabled = False
        DisebleProfileAllElements
    End If
    InitGDI
    
    'Init BlocksSokets
    BlocksSocketShared.HookForm hWndForm
    Set BlocksSocket1 = New BlocksSocket
    Set Timer1 = New Timer
End Sub

' -----------------------------------------------------------------------------
' SetTexts (Sub), source lines 306-343
' -----------------------------------------------------------------------------
Private Sub SetTexts()
    Frame2.Caption = localization.ActiveLng.Item("RaskroyForm.Frame2.Caption")
    CommandNew.Caption = localization.ActiveLng.Item("RaskroyForm.CommandNew.Caption")
    CommandSave.Caption = localization.ActiveLng.Item("RaskroyForm.CommandSave.Caption")
    CommandDelete.Caption = localization.ActiveLng.Item("RaskroyForm.CommandDelete.Caption")
    MultiPage1.Pages("Page1").Caption = localization.ActiveLng.Item("RaskroyForm.MultiPage1.Pages.Page1.Caption")
    MultiPage1.Pages("Page2").Caption = localization.ActiveLng.Item("RaskroyForm.MultiPage1.Pages.Page2.Caption")
    Frame1.Caption = localization.ActiveLng.Item("RaskroyForm.Frame1.Caption")
    Label3.Caption = localization.ActiveLng.Item("RaskroyForm.Label3.Caption")
    Label4.Caption = localization.ActiveLng.Item("RaskroyForm.Label4.Caption")
    Check2.Caption = localization.ActiveLng.Item("RaskroyForm.Check2.Caption")
    Frame3.Caption = localization.ActiveLng.Item("RaskroyForm.Frame3.Caption")
    Label6.Caption = localization.ActiveLng.Item("RaskroyForm.Label6.Caption")
    Label5.Caption = localization.ActiveLng.Item("RaskroyForm.Label5.Caption")
    Label7.Caption = localization.ActiveLng.Item("RaskroyForm.Label7.Caption")
    Check1.Caption = localization.ActiveLng.Item("RaskroyForm.Check1.Caption")
    Label8.Caption = localization.ActiveLng.Item("RaskroyForm.Label8.Caption")
    Label9.Caption = localization.ActiveLng.Item("RaskroyForm.Label9.Caption")
    
    Check4.Caption = localization.ActiveLng.Item("RaskroyForm.Check4.Caption")
    Check3.Caption = localization.ActiveLng.Item("RaskroyForm.Check3.Caption")
    Check5.Caption = localization.ActiveLng.Item("RaskroyForm.Check5.Caption")
    Check6.Caption = localization.ActiveLng.Item("RaskroyForm.Check6.Caption")
    Check7.Caption = localization.ActiveLng.Item("RaskroyForm.Check7.Caption")
    Check8.Caption = localization.ActiveLng.Item("RaskroyForm.Check8.Caption")
    Check9.Caption = localization.ActiveLng.Item("RaskroyForm.Check9.Caption")
    CheckBox1.Caption = localization.ActiveLng.Item("RaskroyForm.CheckBox1.Caption")
    dntChangeCheckBox1 = True
    CheckBox1.value = RegData1.NestingPlaceOnPage
    dntChangeCheckBox1 = False

    
    Command4.Caption = localization.ActiveLng.Item("RaskroyForm.Command4.Caption")
    Command5.Caption = localization.ActiveLng.Item("RaskroyForm.Command5.Caption")
    Command1.Caption = localization.ActiveLng.Item("RaskroyForm.Command1.Caption")
    
    Combo2.List(0) = localization.ActiveLng.Item("RaskroyForm.withoutTurning")
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 344-352
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()
    On Error GoTo 10
    If FirstRun = False Then
        FirstRun = True
        FirstRunSub
        Timer1.Start 400
    End If
10
End Sub

' -----------------------------------------------------------------------------
' FirstRunSub (Sub), source lines 353-366
' -----------------------------------------------------------------------------
Private Sub FirstRunSub()
    'Первый запуск
    RaskroyForm.Repaint
'    UngroupOneElementGroupsStart 'Delete 1-group elements for PowerClip
    TakeData
    CurveData = PrepareToSend(RCurves()) 'Подготавливает данные к отправке, конвертирует всё в Bytes
    
    If Break = True Then Exit Sub
    
    Label1.Caption = ""
    ProgressBar1.value = 0
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
End Sub

' -----------------------------------------------------------------------------
' InitGDI (Sub), source lines 395-415
' -----------------------------------------------------------------------------
Private Sub InitGDI()
    ' Get hWndForm
    hWndForm = FindWindow(vbNullString, Me.Caption)
    HDCForm = GetDC(hWndForm)
    
    Dim gdipStartupInput As GDIPlusStartupInput
    gdipStartupInput.GdiPlusVersion = 1&
    GdiplusStartup GdipToken, gdipStartupInput, ByVal 0
    GdipCreateFromHDC HDCForm, HGraphics
    GdipCreatePen1 ConvertColor(vbBlack, 100), 1, 2, mPen
    GdipCreatePen1 ConvertColor(vbBlack, 50), 1, 2, LinePen
    GdipCreatePen1 ConvertColor(RGB(0, 0, 150), 100), 1, 2, MaterialPen
    GdipCreatePen1 ConvertColor(RGB(255, 98, 0), 100), 1, 2, PagesPen
    GdipCreatePen1 ConvertColor(RGB(255, 173, 122), 60), 1, 2, FieldsPen
    GdipCreatePen1 ConvertColor(vbWhite, 100), 1, 2, LinesPaperPen
    GdipCreatePen1 ConvertColor(vbBlack, 30), 1, 2, LinesShadowPen
    GdipCreateSolidFill ConvertColor(vbWhite, 100), PaperBrush
    GdipCreateSolidFill ConvertColor(RGB(205, 205, 205), 100), hBrush
    GdipCreateSolidFill ConvertColor(vbBlack, 30), ShadowBrush
    GdipSetSmoothingMode HGraphics, SmoothingModeAntiAlias
End Sub

' -----------------------------------------------------------------------------
' Timer1_OnTimer (Sub), source lines 416-433
' -----------------------------------------------------------------------------
Private Sub Timer1_OnTimer()
    If Connected Then
        With BlocksSocket1
            If .Statistic_PING < 1000 Then
                If .Statistic_PING = LastShowedPing Then .Statistic_PING = .Statistic_PING - 1
                Label11.Caption = "Ping " & .Statistic_PING & " ms"
                LastShowedPing = .Statistic_PING
                .PING
                .Statistic_PING = .Statistic_PING + 400
            Else
                Label11.Caption = ""
                .PING
            End If
        End With
    Else
        Label11.Caption = ""
    End If
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnConnectionProcessIsStarted (Sub), source lines 434-436
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnConnectionProcessIsStarted()
    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.connectTo") & " " & BlocksSocket1.Address & ".."
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnConnect (Sub), source lines 437-487
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnConnect(client As BlocksSocket)
    Label11.Visible = True
    Connected = True

    'Send Head
    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 60
        .bPut_Long Commands.HEAD_Raskroy_20220421
        
        .bPut_Long UserData.PlotCalcVersion
        .bPut_Long SessionID
        If Application.VersionMajor = 24 And Application.VersionMinor >= 3 Then
            'Corel 2023
            .bPut_Long Application.VersionMajor + 1999
        Else
            If Application.VersionMajor >= 25 Then
                'Corel 2024 +
                .bPut_Long Application.VersionMajor + 1999
            Else
                .bPut_Long Application.VersionMajor + 1998
            End If
        End If
'        application.corelv
        Dim lcode As String * 2
        lcode = UserData.LangCode
        .bPut_String lcode
        Dim code As String * 20
        code = Trim(UserData.CodeOfActivation)
        If Len(code) <> 20 Then code = "--------------------"
        .bPut_String code
        
        BlocksSocket1.Send .bytes
    End With
    If SessionID <> 0 Then
        ProtocolStatus = RepairProtocolStatus
        Label1.Caption = LastCaption
    Else
        ProtocolStatus = WaitingForActivationStatus
        Label1.Caption = ""
    End If
    
    If ActivationFormIsShow Then
        
        ActivationForm.Text1.Enabled = True
        If Len(ActivationForm.Text1.text) = 20 Then
            ActivationForm.Command1.Enabled = True
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnRecive (Sub), source lines 488-624
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnRecive(bytes() As Byte, client As BlocksSocket)
    On Error GoTo 10
    If UBound(bytes) > 2 Then
        Dim buffer1 As New buffer
        Dim command As Commands
        buffer1.InterPlatformNetConverter = True
        buffer1.SetBytes bytes
        command = buffer1.bGet_Long
        
        If command = SetSessionID Then
            SessionID = buffer1.bGet_Long
            Exit Sub
        ElseIf command = ErrorDoubleUseCode Then
            Label1.Caption = "Error Double Use Code " & UserData.CodeOfActivation
            Error = True
            SessionID = 0
            BlocksSocket1.Disconnect
            Exit Sub
        ElseIf command = ErrorSessionNotFound Then
            Label1.Caption = "Session is not found"
            Error = True
            SessionID = 0
            BlocksSocket1.Disconnect
            Exit Sub
        ElseIf command = CloseSession Then
            Label1.Caption = "Session is closed by the PlotCalc server"
            Error = True
            SessionID = 0
            BlocksSocket1.Disconnect
            Exit Sub
        End If
        
        Select Case ProtocolStatus
            Case ProtocolStatusEnm.WaitingForActivationStatus
                If command = TimeBeforeFreeWork Then
                    Dim t As Long
                    t = buffer1.bGet_Long
                    If t = 0 Then
                        Plotf = False
                    Else
                        Plotf = True
                        Time_for_work = t
                        ActivationFormIsShow = True
                        Set ActivationForm.ParentForm = Me
                        ActivationForm.Show 1
                        ActivationFormIsShow = False
                    End If
                    ProtocolStatus = Ready
                    EnableCommand1
                    If Command1.Enabled = True Then Command1.SetFocus
                ElseIf command = ActivationSuccess Then 'Активация успешна
                    MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox1"), vbInformation, "PlotCalc"
                    Plotf = False
                    Unload ActivationForm
                    
                ElseIf command = ActivationCodeNotExists Then 'Не сущетвует такого кода
                    MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox2"), vbCritical, "PlotCalc"
                    
                ElseIf command = ActivationCodeTimeOut Then 'Время действия лицензии истекло
                    MsgBox localization.ActiveLng.Item("ActivationForm.MsgBox3"), vbInformation, "PlotCalc"
                Else
                    Label1.Caption = "Unknown command"

                End If
                Timer1.Enabled = True
            Case ProtocolStatusEnm.WaitingDataSentResult
                If command = ServerReceivedAllDataFlag Then
                    Sending_Data = True
                    Command1.Enabled = True
                    Command1.SetFocus
                    ProtocolStatus = WaitingNestingResult
                    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.WaitingServerResponse")
                    ProgressBar1.value = 0
                    StartNesting
                End If
            Case ProtocolStatusEnm.WaitingNestingResult
                If command = ProgressPercent Then
                    Dim Percent As Long
                    Percent = buffer1.bGet_Long
                    ProgressBar1.value = Percent
                    If ImproveOn Then
                        pp = buffer1.bGet_Long
                        Label10.Caption = localization.ActiveLng.Item("RaskroyForm.Rebuild") & " " & pp & "   " & localization.ActiveLng.Item("RaskroyForm.Improved") & " " & Improved & MmPlus
                    End If
                ElseIf command = result Then
                    If Not ImproveOn Then Label1.Caption = localization.ActiveLng.Item("RaskroyForm.GettingResult")
                    ProgressBar1.value = 0
                    CountResults = CountResults + 1
                    ReDim Preserve Work_Result(1 To CountResults)
                    Work_Result(CountResults) = Get_Work_Result(buffer1) 'Забираем результат
                    SelectedResult = CountResults
                    
                    'Send more
                    If ImproveOn Then
                        Improved = Improved + 1
                        Label10.Caption = localization.ActiveLng.Item("RaskroyForm.Rebuild") & " " & pp & "   " & localization.ActiveLng.Item("RaskroyForm.Improved") & " " & Improved & MmPlus
                    Else
                        Label1.Caption = localization.ActiveLng.Item("RaskroyForm.ResultObtained") & "   (" & Int(Work_Result(SelectedResult).maximumY * 25.4) & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
                        LastMaxY = Work_Result(SelectedResult).maximumY
                        Improved = 0
                        Command1.Caption = localization.ActiveLng.Item("RaskroyForm.ToImprove")
                        Command1.BackColor = RGB(147, 244, 186)
                        ProgressBar1.GreenColorOn = True
                        If Work_Result(SelectedResult).Errors Then Command1.Enabled = False
                        ImproveOn = True
                        EnableAllButtons
                        Processing_Data = False
                        ProtocolStatus = Ready
                    End If
                    'Теперь нарисовать результат
                    Dim wrtemp As Work_Result
                    wrtemp = Work_Result(SelectedResult)
                    DrawWarkResult wrtemp
                    
                ElseIf command = NoMoreCombinations Then
                    Command1.Enabled = False
                    Command1.Caption = localization.ActiveLng.Item("RaskroyForm.ToImprove")
                    Command1.BackColor = VBA.SystemColorConstants.vbButtonFace
                    ProgressBar1.GreenColorOn = False
                    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.impossibleToImproveMore") ' & MmPlus
                    EnableAllButtons
                    Processing_Data = False
                    ProtocolStatus = Ready
                    ProgressBar1.value = 0
                ElseIf command = ErrorContourCreation Then
                    Label1.Caption = "Error Contour Creation"
                ElseIf command = ResultAsBMP Then
                    Label1.Caption = "Result as BMP. Update software"
                    Error = True
                    SessionID = 0
                    BlocksSocket1.Disconnect
                End If
            Case ProtocolStatusEnm.Ready
        End Select
    End If
10
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnDeliveryReporting (Sub), source lines 625-628
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnDeliveryReporting(BytesDelivered As Long, BytesSize As Long, client As BlocksSocket)
    Label1.Caption = localization.ActiveLng.Item("RaskroyForm.SendingDataServer") & "   (" & BytesDelivered \ 1000 & "/" & BytesSize \ 1000 & " KB)"
    ProgressBar1.value = (CDbl(BytesDelivered) / CDbl(BytesSize)) * 10000
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnDisconnect (Sub), source lines 635-662
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnDisconnect(client As BlocksSocket)
    Connected = False
    If ActivationFormIsShow Then
        ActivationForm.Command1.Enabled = False
        ActivationForm.Text1.Enabled = False
    End If
    
    If SessionID <> 0 Then
        'Repair Session
        RepairProtocolStatus = ProtocolStatus
        LastCaption = Label1.Caption
        Set BlocksSocket1 = Nothing
        Set BlocksSocket1 = New BlocksSocket
        BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
    Else
        'End
        Label11.Visible = False
        ProgressBar1.value = 0
        Command1.Enabled = False
        Timer1.Enabled = False
        ProtocolStatus = NotConnected
        If Not Error Then
'            Label1.Font.Bold = True
'            Label1.ForeColor = RGB(240, 68, 77)
            Label1.Caption = localization.ActiveLng.Item("RaskroyForm.NoConnectionServer") & " (" & BlocksSocket1.Address & ")"
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' Command1_Click (Sub), source lines 663-744
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' MmPlus (Function), source lines 745-751
' -----------------------------------------------------------------------------
Private Function MmPlus() As String
    On Error GoTo e:
    Dim mm As Long
    mm = Int((LastMaxY - Work_Result(CountResults).maximumY) * 25.4)
    If mm > 0 Then MmPlus = " (+" & mm & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
e:
End Function

' -----------------------------------------------------------------------------
' StartNesting (Sub), source lines 752-821
' -----------------------------------------------------------------------------
Private Sub StartNesting()
    'Передать параметры и начать процесс размещения
    Dim To4nost As Long:                    To4nost = 20 'Количество полос
    Dim GradusCount2 As Long
    Dim Vlogen As Boolean

    SingMaterialWith = DelTochka(Text4) * 1000
    SingMaterialHeight = DelTochka(Text5) * 1000
    SingPlotterSheetWith = (DelTochka(Text8) - (DelTochka(Text13) + DelTochka(Text14))) * 1000
    SingPlotterSheetHeight = (DelTochka(Text7) - (DelTochka(Text12) + DelTochka(Text15))) * 1000
    SingLeftField = DelTochka(Text13) * 1000
    SingRightField = DelTochka(Text14) * 1000
    SingUpField = DelTochka(Text12) * 1000
    SingDownField = DelTochka(Text15) * 1000
    Otstup = DelTochka(Text18) / 2 * 1000

    Select Case Combo2.ListIndex
        Case 0
            GradusCount2 = 1
        Case 1
            GradusCount2 = 2
        Case 2
            GradusCount2 = 4
        Case 3
            GradusCount2 = 8
        Case Else
            GradusCount2 = 32  'Кол-во градусов
    End Select
    
    If Check1.value Then
        Vlogen = False 'Разрешить вложение
    Else
        Vlogen = True 'Запретить вложение
    End If
    
    If Check2.value Then
        Vnutri_Figuri = True 'Внутри самой Большой фигуры
        SingMaterialWith = AlfaWidth * 25.4 * 1000 + 10
        SingMaterialHeight = AlfaHeight * 25.4 * 1000 + 10
        SingPlotterSheetWith = AlfaWidth * 25.4 * 1000 + 10
        SingPlotterSheetHeight = AlfaHeight * 25.4 * 1000 + 10
        SingLeftField = 0
        SingRightField = 0
        SingUpField = 0
        SingDownField = 0
    Else
        Vnutri_Figuri = False 'Внутри самой Большой фигуры
    End If
        
    Dim buffer1 As New buffer
    With buffer1
        .InterPlatformNetConverter = True
        .size = 52
        .bPut_Long Commands.Place
        .bPut_Long GradusCount2
        .bPut_Long To4nost
        .bPut_Long Otstup
        .bPut_Long SingMaterialWith
        .bPut_Long SingMaterialHeight
        .bPut_Long SingPlotterSheetWith
        .bPut_Long SingPlotterSheetHeight
        .bPut_Long SingLeftField
        .bPut_Long SingRightField
        .bPut_Long SingUpField
        .bPut_Long SingDownField
        .bPut_Boolean Vnutri_Figuri
        .bPut_Boolean Vlogen
    End With
    BlocksSocket1.Send buffer1.bytes
End Sub

' -----------------------------------------------------------------------------
' DrawWarkResult (Sub), source lines 822-921
' -----------------------------------------------------------------------------
Private Sub DrawWarkResult(ByRef Work_Result As Work_Result)
    DrawWarkResultProcessing = True
    On Error Resume Next
    Dim ks As Double
    Dim dx As Double, dy As Double
    Dim oblastH As Double, oblastW As Double
    Dim minusH As Double, minusW As Double
    Dim horisontal As Boolean 'Горихонталь или вертикаль
    Dim plusY As Double, plusX As Double
    minusH = 0: minusW = 0
    plusX = 0: plusY = 0
    If Work_Result.SingMaterialHeight < Work_Result.SingMaterialWith Then
        oblastH = Work_Result.SingMaterialHeight * Work_Result.CountMaterials
        oblastW = Work_Result.SingMaterialWith
        minusH = 5 * Work_Result.CountMaterials - 5
        horisontal = False
    Else
        oblastH = Work_Result.SingMaterialHeight
        oblastW = Work_Result.SingMaterialWith * Work_Result.CountMaterials
        minusW = 5 * Work_Result.CountMaterials - 5
        horisontal = True
    End If
    Cls
    If oblastW / WindowWidth > oblastH / WindowHeigth Then
        ks = (WindowWidth - 15 - minusW) / oblastW
    Else
        ks = (WindowHeigth - 20 - minusH) / oblastH
    End If
    dx = 6
    dy = 8
    Dim q As Long
    'Теперь нарисовать фигуры ===================================
    Dim q2 As Long
    Dim q3 As Long
    Dim q4 As Long
    Dim xx1 As Double, w As Long
    Dim yy1 As Double, h As Long
    Dim hh As Long, hh2 As Long
    w = Work_Result.SingMaterialWith * ks
    h = Work_Result.SingMaterialHeight * ks
    plusX = 0
    plusY = 0
    For q = 1 To Work_Result.CountMaterials
        With Work_Result.Materials(q)
            If horisontal Then
                plusX = (q - 1) * Work_Result.SingMaterialWith + (q - 1) * 5 / ks
                plusY = Work_Result.SingDownField
            Else
                plusX = 0
                plusY = (q - 1) * Work_Result.SingMaterialHeight + Work_Result.SingDownField + (q - 1) * 5 / ks
            End If
            xx1 = dx + plusX * ks
            yy1 = WindowHeigth - dy - plusY * ks + Work_Result.SingDownField * ks - Work_Result.SingMaterialHeight * ks
            If q = 1 And Work_Result.Vnutri_Figuri = True Then
                DrawFigList ks
            Else
                FillRect ShadowBrush, xx1 + 2, yy1 + 2, w, h
                FillRect PaperBrush, xx1, yy1, w, h
                DrawRect LinePen, xx1, yy1, w, h
            End If
            For q2 = 1 To .CountRLines
                With .RLines(q2)
                    If horisontal Then
                        plusX = (q - 1) * Work_Result.SingMaterialWith + Work_Result.SingLeftField + (q - 1) * 5 / ks
                    Else
                        plusX = Work_Result.SingLeftField
                    End If
                    hh = (Work_Result.SingUpField + .Height + Work_Result.SingDownField) * ks
                    hh2 = .Height * ks
                    For q3 = 1 To .CountRSheets
                        With .RSheets(q3)
                            For q4 = 1 To .CountRFigures
                                With .RFigures(q4)
                                    DrawFigure RCurves(.Q_figure), .x + plusX - Work_Result.Otstup * 2, .y + plusY - Work_Result.Otstup * 2, .gradus, ks
                                End With
                                DoEvents
                                If Not DrawWarkResultProcessing Then Exit Sub
                            Next
                            'Нарисовать границы области
                            If .CountRFigures > 0 And Work_Result.Vnutri_Figuri = False Then
                                xx1 = dx + plusX * ks - Work_Result.SingLeftField * ks
                                yy1 = WindowHeigth - dy - plusY * ks - hh + Work_Result.SingDownField * ks
                                DrawRect FieldsPen, xx1 + Work_Result.SingLeftField * ks, yy1 + Work_Result.SingUpField * ks, .Width * ks, hh2
                                DrawRect PagesPen, xx1, yy1, (.Width + Work_Result.SingLeftField + Work_Result.SingRightField) * ks, hh
                            End If
                            plusX = plusX + Work_Result.SingRightField + .Width + Work_Result.SingLeftField
                        End With
                    Next
                    plusY = plusY + Work_Result.SingDownField + .Height + Work_Result.SingUpField
                End With
            Next
        End With
    Next
    If Work_Result.Errors = True Then
        Dim str As String
        str = localization.ActiveLng.Item("RaskroyForm.NOTE_SomeCharactersNotPossible")
        TextOut HDCForm, shiftX + 10, shiftY + WindowHeigth - 30, str, Len(str)
    End If
    DrawWarkResultProcessing = False
End Sub

' -----------------------------------------------------------------------------
' DrawFigure (Sub), source lines 922-1034
' -----------------------------------------------------------------------------
Private Sub DrawFigure(ByRef Figure As RCurve, ByVal x As Double, ByVal y As Double, ByVal gradus As Double, ByVal ks As Double)
    'Создаем просто массив линий с преодразованием безъе в множество прямых линий
    Dim q As Long, q2 As Long
    Dim sNX As Double, sNY As Double
    Dim eNX As Double, eNY As Double
    Dim SCPx As Double, SCPy As Double
    Dim ECPx As Double, ECPy As Double
    Dim minY As Double, maxY As Double
    Dim minX As Double, maxX As Double
    Dim aa As Long
    Dim sheetCurve As PLCurve
    Dim maxRecurse As Long
    Dim a As Double, a1 As Double, a2 As Double
    Dim da1 As Double, da2 As Double
    minX = 999999999999#
    minY = 999999999999#
    sheetCurve.CountPaths = UBound(Figure.RPaths)
    ReDim sheetCurve.Paths(1 To sheetCurve.CountPaths)
    With Figure
        For q = 1 To UBound(.RPaths)
            aa = 1
            sNX = .RPaths(q).RSegments(1).SNPx * .MaxLenght
            sNY = .RPaths(q).RSegments(1).SNPy * .MaxLenght
            ReDim sheetCurve.Paths(q).segment(1 To 1)
            sheetCurve.Paths(q).segment(1).Xpoint = sNX
            sheetCurve.Paths(q).segment(1).Ypoint = sNY
            With .RPaths(q)
                For q2 = 1 To UBound(.RSegments)
                    SCPx = .RSegments(q2).SCPx * Figure.MaxLenght
                    SCPy = .RSegments(q2).SCPy * Figure.MaxLenght
                    ECPx = .RSegments(q2).ECPx * Figure.MaxLenght
                    ECPy = .RSegments(q2).ECPy * Figure.MaxLenght
                    If q2 = UBound(.RSegments) Then
                        eNX = sheetCurve.Paths(q).segment(1).Xpoint
                        eNY = sheetCurve.Paths(q).segment(1).Ypoint
                    Else
                        eNX = .RSegments(q2 + 1).SNPx * Figure.MaxLenght
                        eNY = .RSegments(q2 + 1).SNPy * Figure.MaxLenght
                    End If
                    a1 = Angel(sNX, sNY, SCPx, SCPy)
                    a = Angel(SCPx, SCPy, ECPx, ECPy)
                    a2 = Angel(ECPx, ECPy, eNX, eNY)
                    da1 = Abs(DeltaAngle(a, a1))
                    da2 = Abs(DeltaAngle(a2, a))
                    If SCPx = sNX And SCPy = sNY Then da1 = 0
                    If ECPx = eNX And ECPy = eNY Then da2 = 0
                    a = da1 + da2
                    If a < 0.01 Then
                        maxRecurse = 0
                    ElseIf a < 0.05 Then
                        maxRecurse = 1
                    ElseIf a < 0.26 Then
                        maxRecurse = 2
                    ElseIf a < 0.9 Then
                        maxRecurse = 3
                    ElseIf a < 2.2 Then
                        maxRecurse = 4
                    Else
                        maxRecurse = 5
                    End If
                    recursive_bezier sNX, sNY, SCPx, SCPy, ECPx, ECPy, eNX, eNY, 1, sheetCurve.Paths(q), aa, maxRecurse, maxX, minX, maxY, minY
                    sNX = eNX
                    sNY = eNY
                Next
            End With
        Next
    End With
    
    Dim LinesPaths() As LinesPath
    ReDim LinesPaths(1 To UBound(Figure.RPaths))

    maxX = -9999999999#: maxY = -9999999999#
    Dim tX As Double, tY As Double

    With sheetCurve
        For q = 1 To UBound(.Paths)
            With .Paths(q)
                For q2 = 1 To UBound(.segment)
                    With .segment(q2)
                        LinesPaths(q).CountSegment = LinesPaths(q).CountSegment + 1
                        ReDim Preserve LinesPaths(q).LineSegments(1 To LinesPaths(q).CountSegment)
                        tX = .Xpoint
                        tY = .Ypoint
                        With LinesPaths(q).LineSegments(q2)
                            .x = tX * Cos(gradus) + tY * Sin(gradus)
                            .y = tY * Cos(gradus) - tX * Sin(gradus)
                            If .x > maxX Then maxX = .x
                            If .y > maxY Then maxY = .y
                        End With
                    End With
                Next
            End With
        Next
    End With
    '=========================================================
    Dim dx As Long, dy As Long
    dx = 6
    dy = 8
    Dim x1 As Long, y1 As Long
    Dim x2 As Long, y2 As Long
    
    For q = 1 To UBound(LinesPaths)
        With LinesPaths(q)
            For q2 = 1 To .CountSegment - 1
                x1 = (.LineSegments(q2).x - maxX) * ks + x * ks
                y1 = (.LineSegments(q2).y - maxY) * ks + y * ks
                x2 = (.LineSegments(q2 + 1).x - maxX) * ks + x * ks
                y2 = (.LineSegments(q2 + 1).y - maxY) * ks + y * ks
                DrawLine LinePen, dx + x1, WindowHeigth - y1 - dy, dx + x2, WindowHeigth - y2 - dy
            Next
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' DrawFigList (Sub), source lines 1035-1153
' -----------------------------------------------------------------------------
Private Sub DrawFigList(ByRef ks As Double)
    'NumberOfMaxFigure
    On Error Resume Next
    Dim q As Long, q2 As Long
    Dim sNX As Double, sNY As Double
    Dim eNX As Double, eNY As Double
    Dim SCPx As Double, SCPy As Double
    Dim ECPx As Double, ECPy As Double
    Dim minY As Double, maxY As Double
    Dim minX As Double, maxX As Double
    Dim aa As Long
    Dim sheetCurve As PLCurve
    Dim maxRecurse As Long
    Dim a As Double, a1 As Double, a2 As Double
    Dim da1 As Double, da2 As Double
    minX = 999999999999#
    minY = 999999999999#
    sheetCurve.CountPaths = UBound(RCurves(NumberOfMaxFigure).RPaths)
    ReDim sheetCurve.Paths(1 To sheetCurve.CountPaths)
    With RCurves(NumberOfMaxFigure)
        For q = 1 To UBound(.RPaths)
            aa = 1
            sNX = .RPaths(q).RSegments(1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
            sNY = .RPaths(q).RSegments(1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
            ReDim sheetCurve.Paths(q).segment(1 To 1)
            sheetCurve.Paths(q).segment(1).Xpoint = sNX
            sheetCurve.Paths(q).segment(1).Ypoint = sNY
            With .RPaths(q)
                For q2 = 1 To UBound(.RSegments)
                    SCPx = .RSegments(q2).SCPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    SCPy = .RSegments(q2).SCPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPx = .RSegments(q2).ECPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPy = .RSegments(q2).ECPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    If q2 = UBound(.RSegments) Then
                        eNX = sheetCurve.Paths(q).segment(1).Xpoint
                        eNY = sheetCurve.Paths(q).segment(1).Ypoint
                    Else
                        eNX = .RSegments(q2 + 1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                        eNY = .RSegments(q2 + 1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    End If
                    a1 = Angel(sNX, sNY, SCPx, SCPy)
                    a = Angel(SCPx, SCPy, ECPx, ECPy)
                    a2 = Angel(ECPx, ECPy, eNX, eNY)
                    da1 = Abs(DeltaAngle(a, a1))
                    da2 = Abs(DeltaAngle(a2, a))
                    If SCPx = sNX And SCPy = sNY Then da1 = 0
                    If ECPx = eNX And ECPy = eNY Then da2 = 0
                    a = da1 + da2
                    If a < 0.01 Then
                        maxRecurse = 0
                    ElseIf a < 0.05 Then
                        maxRecurse = 1
                    ElseIf a < 0.26 Then
                        maxRecurse = 2
                    ElseIf a < 0.9 Then
                        maxRecurse = 3
                    ElseIf a < 2.2 Then
                        maxRecurse = 4
                    Else
                        maxRecurse = 5
                    End If
                    recursive_bezier sNX, sNY, SCPx, SCPy, ECPx, ECPy, eNX, eNY, 1, sheetCurve.Paths(q), aa, maxRecurse, maxX, minX, maxY, minY
                    sNX = eNX
                    sNY = eNY
                Next
            End With
        Next
    End With
    'Теперь нарисовать лист
    Dim dx As Long, dy As Long
    dx = 6
    dy = 8
    For q = 1 To UBound(sheetCurve.Paths)
        With sheetCurve.Paths(q)
            For q2 = 1 To UBound(.segment)
                .segment(q2).Ypoint = -dy + WindowHeigth - .segment(q2).Ypoint + minY
                .segment(q2).Xpoint = dx + .segment(q2).Xpoint - minX
            Next
        End With
    Next
    Dim lines() As LinesString
    ReDim lines(1 To WindowHeigth) As LinesString
    Dim y As Long
    Dim x1 As Double, x2 As Double
    Dim y1 As Double, y2 As Double
    
    For y = WindowHeigth - (maxY - minY) - dy To WindowHeigth
        For q = 1 To UBound(sheetCurve.Paths)
            With sheetCurve.Paths(q)
                For q2 = 1 To UBound(.segment) - 1
                    y1 = .segment(q2).Ypoint
                    y2 = .segment(q2 + 1).Ypoint
                    If y > y1 And y > y2 Then GoTo nx
                    If y < y1 And y < y2 Then GoTo nx
                    If y2 = y1 Then GoTo nx
                    x1 = .segment(q2).Xpoint
                    x2 = .segment(q2 + 1).Xpoint
                    lines(y).CountX = lines(y).CountX + 1
                    ReDim Preserve lines(y).x(1 To lines(y).CountX)
                    lines(y).x(lines(y).CountX) = (x2 - x1) * (y - y1) / (y2 - y1) + x1
nx:
                Next
            End With
        Next
    Next
    For q = 1 To UBound(lines)
        With lines(q)
            SortData .x()
        End With
    Next
    For q = 1 To UBound(lines)
        With lines(q)
            For q2 = 1 To .CountX Step 2
                DrawLine LinesPaperPen, .x(q2), q, .x(q2 + 1), q
                DrawLine LinesShadowPen, .x(q2) + 2, q + 2, .x(q2 + 1) + 2, q + 2
            Next
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' Get_Work_Result (Function), source lines 1154-1207
' -----------------------------------------------------------------------------
Private Function Get_Work_Result(ByRef buffer1 As buffer) As Work_Result
    On Error GoTo 50
    Dim q As Long, q2 As Long, q3 As Long, q4 As Long
    Dim errFlag As Long
    With Get_Work_Result
        .CountMaterials = buffer1.bGet_Long
        If buffer1.bGet_Long = 1 Then .Errors = True
        .maximumY = CDbl(buffer1.bGet_Long) / 25.4 / 1000
        ReDim .Materials(1 To .CountMaterials)
        For q = 1 To .CountMaterials
            With .Materials(q)
                .CountRLines = buffer1.bGet_Long
                ReDim .RLines(1 To .CountRLines)
                For q2 = 1 To .CountRLines
                    With .RLines(q2)
                        .CountRSheets = buffer1.bGet_Long
                        .Height = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                        ReDim .RSheets(1 To .CountRSheets)
                        For q3 = 1 To .CountRSheets
                            With .RSheets(q3)
                                .CountRFigures = buffer1.bGet_Long
                                .Width = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                                If .CountRFigures > 0 Then
                                ReDim .RFigures(1 To .CountRFigures)
                                    For q4 = 1 To .CountRFigures
                                        With .RFigures(q4)
                                            .gradus = CDbl(buffer1.bGet_Long) / 100000000
                                            .Q_figure = buffer1.bGet_Long
                                            .x = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                                            .y = CDbl(buffer1.bGet_Long) / 25.4 / 1000
                                        End With
                                    Next
                                End If
                            End With
                        Next
                    End With
                Next
            End With
        Next
    End With
50
    With Get_Work_Result
        .SingMaterialWith = (SingMaterialWith / 1000) / 25.4
        .SingMaterialHeight = (SingMaterialHeight / 1000) / 25.4
        .SingPlotterSheetWith = (SingPlotterSheetWith / 1000) / 25.4
        .SingPlotterSheetHeight = (SingPlotterSheetHeight / 1000) / 25.4
        .SingLeftField = (SingLeftField / 1000) / 25.4
        .SingRightField = (SingRightField / 1000) / 25.4
        .SingUpField = (SingUpField / 1000) / 25.4
        .SingDownField = (SingDownField / 1000) / 25.4
        .Otstup = (Otstup / 1000) / 25.4
        .Vnutri_Figuri = Vnutri_Figuri
    End With
End Function

' -----------------------------------------------------------------------------
' Binary64BitToDouble (Function), source lines 1208-1219
' -----------------------------------------------------------------------------
Private Function Binary64BitToDouble(ByRef BinaryString As String) As Double
    Dim bytArray(7) As Byte
    bytArray(7) = Asc(Mid(BinaryString, 8, 1))
    bytArray(6) = Asc(Mid(BinaryString, 7, 1))
    bytArray(5) = Asc(Mid(BinaryString, 6, 1))
    bytArray(4) = Asc(Mid(BinaryString, 5, 1))
    bytArray(3) = Asc(Mid(BinaryString, 4, 1))
    bytArray(2) = Asc(Mid(BinaryString, 3, 1))
    bytArray(1) = Asc(Mid(BinaryString, 2, 1))
    bytArray(0) = Asc(Mid(BinaryString, 1, 1))
    CopyMemory Binary64BitToDouble, bytArray(0), 8
End Function

' -----------------------------------------------------------------------------
' strTo16 (Function), source lines 1220-1222
' -----------------------------------------------------------------------------
Private Function strTo16(ByRef str) As Long
    strTo16 = CDbl(Asc(Mid(str, 1, 2))) * 256 + Asc(Mid(str, 2, 2))
End Function

' -----------------------------------------------------------------------------
' Convert64ToSTR (Function), source lines 1223-1227
' -----------------------------------------------------------------------------
Private Function Convert64ToSTR(ByRef s As Double) As String
    Dim bytArray(7) As Byte
    CopyMemory bytArray(0), s, 8
    Convert64ToSTR = Chr(bytArray(0)) & Chr(bytArray(1)) & Chr(bytArray(2)) & Chr(bytArray(3)) & Chr(bytArray(4)) & Chr(bytArray(5)) & Chr(bytArray(6)) & Chr(bytArray(7))
End Function

' -----------------------------------------------------------------------------
' ConvertLongToBytes (Sub), source lines 1228-1238
' -----------------------------------------------------------------------------
Private Sub ConvertLongToBytes(ByRef bytePointer As Byte, ByVal s As Long)
    Dim b(3) As Byte
    b(0) = s \ 16777216
    s = s - 16777216 * CLng(b(0))
    b(1) = s \ 65536
    s = s - 65536 * CLng(b(1))
    b(2) = s \ 256
    s = s - 256 * CLng(b(2))
    b(3) = s
    CopyMemory bytePointer, b(0), 4
End Sub

' -----------------------------------------------------------------------------
' Convert64ToBytes (Sub), source lines 1239-1241
' -----------------------------------------------------------------------------
Private Sub Convert64ToBytes(ByRef bytePointer As Byte, ByVal s As Double)
    CopyMemory bytePointer, s, 8
End Sub

' -----------------------------------------------------------------------------
' ConvertLongToSTR (Function), source lines 1243-1250
' -----------------------------------------------------------------------------
Private Function ConvertLongToSTR(ByRef s As Long) As String
    Dim aa As Byte, bb As Byte, cc As Byte, dd As Byte
    aa = Int(s / 16777216)
    bb = Int((s - CDbl(aa) * 16777216) / 65536)
    cc = Int((s - CDbl(aa) * 16777216 - CDbl(bb) * 65536) / 256)
    dd = Int((s - CDbl(aa) * 16777216 - CDbl(bb) * 65536 - CDbl(cc) * 256))
    ConvertLongToSTR = Chr(aa) & Chr(bb) & Chr(cc) & Chr(dd)
End Function

' -----------------------------------------------------------------------------
' DelTochka (Function), source lines 1251-1253
' -----------------------------------------------------------------------------
Private Function DelTochka(ByRef txt As String) As Double
    DelTochka = val(Replace$(Replace$(txt, ",", "."), "-", ""))
End Function

' -----------------------------------------------------------------------------
' TakeData (Sub), source lines 1254-1288
' -----------------------------------------------------------------------------
Private Sub TakeData()
    On Error GoTo Err:
    Dim q1 As Long
    Dim qq As Long
    Dim CurveFromBitmap As curve

    With VGCore.ActiveSelection
        ReDim RCurves(1 To .shapes.count)
        ReDim Du(1 To .shapes.count)
        ReDim RerotateCrv(1 To .shapes.count)
        For q1 = 1 To .shapes.count
            ProgressBar1.value = Int(q1 / .shapes.count * 10000)
            If Break = True Then Exit Sub
            DoEvents
    
            If Not TestShapeForUsedShapes(.shapes(q1).StaticId, Du(), qq) Then
                qq = qq + 1
                Set Du(qq) = .shapes(q1).GetLinkedShapes(cdrLinkAllConnections)
                If Du(qq).count = 0 Then Du(qq).Add .shapes(q1)
                Set CurveFromBitmap = VGCore.ActiveDocument.CreateCurve
                Set RerotateCrv(qq) = CreateCurveFromShapes(Du(qq).shapes, CurveFromBitmap)
                If RerotateCrv(qq).Subpaths.count = 0 Then
                     Set RerotateCrv(qq) = RerotateCrv(qq).WeldWith(CurveFromBitmap)
                End If
                RCurves(qq) = TakeRCurve(RerotateCrv(qq))
            End If
        Next
        If qq < .shapes.count Then
            ReDim Preserve RCurves(1 To qq)
            ReDim Preserve Du(1 To qq)
            ReDim Preserve RerotateCrv(1 To qq)
        End If
    End With
Err:
End Sub

' -----------------------------------------------------------------------------
' TestShapeForUsedShapes (Function), source lines 1290-1302
' -----------------------------------------------------------------------------
Private Function TestShapeForUsedShapes(StaticId As Long, ByRef sr() As ShapeRange, qq As Long) As Boolean
    Dim q As Long, q2 As Long
    For q = 1 To qq
        With sr(q)
            For q2 = 1 To .shapes.count
                If .shapes(q2).StaticId = StaticId Then
                    TestShapeForUsedShapes = True
                    Exit Function
                End If
            Next
        End With
    Next
End Function

' -----------------------------------------------------------------------------
' CreateCurveFromShapes (Function), source lines 1303-1336
' -----------------------------------------------------------------------------
Private Function CreateCurveFromShapes(shapes As VGCore.shapes, CurveFromBitmap As curve) As curve
    Set CreateCurveFromShapes = VGCore.ActiveDocument.CreateCurve
    Dim q As Long
    For q = 1 To shapes.count
         Select Case shapes(q).Type
            Case cdrShapeType.cdrCurveShape
                Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CloseCurve(shapes(q).curve))
            Case cdrShapeType.cdrGroupShape, cdrShapeType.cdrBlendGroupShape, cdrShapeType.cdrContourGroupShape, cdrArtisticMediaGroupShape
                Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CreateCurveFromShapes(shapes(q).shapes, CurveFromBitmap))
            Case cdrShapeType.cdrBitmapShape
                If shapes(q).Bitmap.Cropped Then
                    Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CloseCurve(shapes(q).Bitmap.CropEnvelope))
                Else
                    With shapes(q)
                        Dim rect2 As curve
                        Set rect2 = VGCore.ActiveDocument.CreateCurve
                        rect2.AppendSubpathRectangle .PositionX, .PositionY, .RightX, .BottomY
                        Set CurveFromBitmap = CurveFromBitmap.WeldWith(rect2)
                    End With
                End If
            Case Else
                If TryDisplayCurve(shapes(q)) Then
                    Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CloseCurve(shapes(q).DisplayCurve))
                Else
                    With shapes(q)
                        Dim rect As curve
                        Set rect = VGCore.ActiveDocument.CreateCurve
                        rect.AppendSubpathRectangle .PositionX, .PositionY, .RightX, .BottomY
                        Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(rect)
                    End With
                End If
         End Select
    Next
End Function

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 1337-1344
' -----------------------------------------------------------------------------
Private Function TryDisplayCurve(shape As shape) As Boolean
    On Error GoTo Err
    Dim c As Long
    c = shape.DisplayCurve.Subpaths.count
    TryDisplayCurve = True
    Exit Function
Err:
End Function

' -----------------------------------------------------------------------------
' CloseCurve (Function), source lines 1345-1353
' -----------------------------------------------------------------------------
Private Function CloseCurve(crv As curve) As curve
    If crv.closed Then
        Set CloseCurve = crv
    Else
        Set CloseCurve = VGCore.ActiveDocument.CreateCurve
        CloseCurve.AppendCurve crv
        CloseCurve.closed = True
    End If
End Function

' -----------------------------------------------------------------------------
' TakeRCurve (Function), source lines 1354-1512
' -----------------------------------------------------------------------------
Private Function TakeRCurve(curve As VGCore.curve) As RCurve
    Dim q2 As Long, q3 As Long
    Dim minY As Double, maxY As Double
    Dim minX As Double, maxX As Double
    Dim pRline As Boolean
    Dim plusDx As Double, plusDy As Double
    
    TakeRCurve.Width = curve.boundingBox.Width
    TakeRCurve.Height = curve.boundingBox.Height
    Dim c As curve
    
    
    minX = 999999999
    minY = 999999999
    maxX = -999999999
    maxY = -999999999
    If TakeRCurve.Width + TakeRCurve.Height > 0.001 Then
        ReDim TakeRCurve.RPaths(1 To curve.Subpaths.count)
        For q2 = 1 To UBound(TakeRCurve.RPaths)
            With curve.Subpaths(q2)
                'If .Closed = False Then .Closed = True
                ReDim TakeRCurve.RPaths(q2).RSegments(1 To .segments.count)
                pRline = False
                For q3 = 1 To UBound(TakeRCurve.RPaths(q2).RSegments)
                    With .segments(q3)
                        TakeRCurve.RPaths(q2).RSegments(q3).SNPx = .StartNode.PositionX
                        TakeRCurve.RPaths(q2).RSegments(q3).SNPy = .StartNode.PositionY
                        If pRline = True Then
                            TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPx = TakeRCurve.RPaths(q2).RSegments(q3).SNPx
                            TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPy = TakeRCurve.RPaths(q2).RSegments(q3).SNPy
                            pRline = False
                        End If
                        With TakeRCurve.RPaths(q2).RSegments(q3)
                            If .SNPx > maxX Then maxX = .SNPx
                            If .SNPx < minX Then minX = .SNPx
                            If .SNPy > maxY Then maxY = .SNPy
                            If .SNPy < minY Then minY = .SNPy
                        End With
                        If .Type = cdrCurveSegment Then
                            TakeRCurve.RPaths(q2).RSegments(q3).SCPx = .StartingControlPointX
                            TakeRCurve.RPaths(q2).RSegments(q3).SCPy = .StartingControlPointY
                            TakeRCurve.RPaths(q2).RSegments(q3).ECPx = .EndingControlPointX
                            TakeRCurve.RPaths(q2).RSegments(q3).ECPy = .EndingControlPointY
                            With TakeRCurve.RPaths(q2).RSegments(q3)
                                If .SCPx > maxX Then maxX = .SCPx
                                If .SCPx < minX Then minX = .SCPx
                                If .SCPy > maxY Then maxY = .SCPy
                                If .SCPy < minY Then minY = .SCPy
                                If .ECPx > maxX Then maxX = .ECPx
                                If .ECPx < minX Then minX = .ECPx
                                If .ECPy > maxY Then maxY = .ECPy
                                If .ECPy < minY Then minY = .ECPy
                            End With
                        Else
                            With TakeRCurve.RPaths(q2).RSegments(q3)
                                .SCPx = .SNPx
                                .SCPy = .SNPy
                                pRline = True
                            End With
                        End If
                    End With
                Next
                    If pRline = True Then
                        TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPx = TakeRCurve.RPaths(q2).RSegments(1).SNPx
                        TakeRCurve.RPaths(q2).RSegments(q3 - 1).ECPy = TakeRCurve.RPaths(q2).RSegments(1).SNPy
                    End If
            End With
        Next
    Else
        If TakeRCurve.Width + TakeRCurve.Height > 0.001 Then
            maxX = curve.boundingBox.Width + curve.boundingBox.x
            minX = curve.boundingBox.x
            maxY = curve.boundingBox.Height + curve.boundingBox.y
            minY = curve.boundingBox.y
        Else
            maxX = 0.001
            maxY = 0.001
            minX = 0
            minY = 0
        End If
        ReDim TakeRCurve.RPaths(1 To 1)
        ReDim TakeRCurve.RPaths(1).RSegments(1 To 4)
        With TakeRCurve.RPaths(1)
            .RSegments(1).SNPx = minX
            .RSegments(1).SNPy = minY
            .RSegments(1).SCPx = minX
            .RSegments(1).SCPy = minY
            .RSegments(1).ECPx = maxX
            .RSegments(1).ECPy = minY
            
            .RSegments(2).SNPx = maxX
            .RSegments(2).SNPy = minY
            .RSegments(2).SCPx = maxX
            .RSegments(2).SCPy = minY
            .RSegments(2).ECPx = maxX
            .RSegments(2).ECPy = maxY
            
            .RSegments(3).SNPx = maxX
            .RSegments(3).SNPy = maxY
            .RSegments(3).SCPx = maxX
            .RSegments(3).SCPy = maxY
            .RSegments(3).ECPx = minX
            .RSegments(3).ECPy = maxY
            
            .RSegments(4).SNPx = minX
            .RSegments(4).SNPy = maxY
            .RSegments(4).SCPx = minX
            .RSegments(4).SCPy = maxY
            .RSegments(4).ECPx = minX
            .RSegments(4).ECPy = minY
        End With
    End If
    'Тепрь надо перевести в относительную систему координат
    For q2 = 1 To UBound(TakeRCurve.RPaths)
        With TakeRCurve.RPaths(q2)
            For q3 = 1 To UBound(.RSegments())
                With .RSegments(q3)
                    .ECPx = SupRound(.ECPx - minX)
                    .ECPy = SupRound(.ECPy - minY)
                    .SCPx = SupRound(.SCPx - minX)
                    .SCPy = SupRound(.SCPy - minY)
                    .SNPx = SupRound(.SNPx - minX)
                    .SNPy = SupRound(.SNPy - minY)
                End With
            Next
        End With
    Next
    
    maxX = SupRound(maxX - minX)
    minX = 0
    maxY = SupRound(maxY - minY)
    minY = 0
    If (maxX) > (maxY) Then
        TakeRCurve.MaxLenght = maxX
        plusDx = 0
        plusDy = SupRound((TakeRCurve.MaxLenght - maxY) / 2)
    Else
        TakeRCurve.MaxLenght = maxY
        plusDx = SupRound((TakeRCurve.MaxLenght - maxX) / 2)
        plusDy = 0
    End If

    If TakeRCurve.MaxLenght = 0 Then TakeRCurve.MaxLenght = 1
    TakeRCurve.MaxLenght = TakeRCurve.MaxLenght / 65530
    '===================================================================================
    'Перестроить числа
        For q2 = 1 To UBound(TakeRCurve.RPaths)
            For q3 = 1 To UBound(TakeRCurve.RPaths(q2).RSegments)
                With TakeRCurve.RPaths(q2).RSegments(q3)
                    .SNPx = (.SNPx - minX + plusDx) / TakeRCurve.MaxLenght
                    .SNPy = (.SNPy - minY + plusDy) / TakeRCurve.MaxLenght
                    .SCPx = (.SCPx - minX + plusDx) / TakeRCurve.MaxLenght
                    .SCPy = (.SCPy - minY + plusDy) / TakeRCurve.MaxLenght
                    .ECPx = (.ECPx - minX + plusDx) / TakeRCurve.MaxLenght
                    .ECPy = (.ECPy - minY + plusDy) / TakeRCurve.MaxLenght
                End With
            Next
        Next
End Function

' -----------------------------------------------------------------------------
' PrepareToSend (Function), source lines 1514-1637
' -----------------------------------------------------------------------------
Private Function PrepareToSend(ByRef RCurves() As RCurve) As Byte()
    On Error GoTo Err
    Dim q1 As Long, q2 As Long, q3 As Long
    Dim kof As Long
    'Сортируем данные
    Dim q As Long
    Dim Keys() As Long
    Dim wi() As Double
    Dim he() As Double
    ReDim Keys(1 To UBound(RCurves))
    ReDim wi(1 To UBound(RCurves))
    ReDim he(1 To UBound(RCurves))
    For q = 1 To UBound(RCurves)
        Keys(q) = q
        wi(q) = RCurves(q).Width
        he(q) = RCurves(q).Height
    Next
    
    Dim i As Long
    Dim fSwitched As Boolean
    Dim strTempW As Double
    Dim strTempH As Double
    Dim keyTemp As Long
    Do
        fSwitched = False
        For i = 1 To UBound(RCurves) - 1
            If wi(i) * he(i) < wi(i + 1) * he(i + 1) Then 'сортируем по возрастанию
                fSwitched = True
                strTempW = wi(i)
                strTempH = he(i)
                keyTemp = Keys(i)
                
                wi(i) = wi(i + 1)
                he(i) = he(i + 1)
                Keys(i) = Keys(i + 1)
                
                wi(i + 1) = strTempW
                he(i + 1) = strTempH
                Keys(i + 1) = keyTemp
            End If
        Next
    Loop While fSwitched

    Dim keysBuff As New buffer
    keysBuff.InterPlatformNetConverter = True
    keysBuff.size = UBound(Keys) * 4 + 4
    
    keysBuff.bPut_Long UBound(Keys)
    For q = 1 To UBound(Keys)
        keysBuff.bPut_Long Keys(q)
    Next

    
    AlfaWidth = RCurves(Keys(1)).Width
    AlfaHeight = RCurves(Keys(1)).Height
    NumberOfMaxFigure = Keys(1)
    
    '==============================================
    Dim curveBuffers() As buffer
    ReDim curveBuffers(1 To UBound(Keys))
    
    Dim qq2 As Long
    Dim KeysQ1 As Long
    
    For q1 = 1 To UBound(RCurves)
        ProgressBar1.value = Int(q1 / UBound(RCurves) * 10000)
        If Break = True Then Exit Function
        DoEvents
        KeysQ1 = Keys(q1)
        Set curveBuffers(KeysQ1) = New buffer
        curveBuffers(KeysQ1).InterPlatformNetConverter = True
        
        With RCurves(KeysQ1)
            If q1 > 1 Then
                qq2 = Keys(q1 - 1)
            Else
                GoTo e
            End If

            If Odinacovo(RCurves(KeysQ1), RCurves(qq2)) Then
                curveBuffers(KeysQ1).size = 4
                curveBuffers(KeysQ1).bPut_Long 1
            Else
e:
                'А иначе если фигура нетакая же как и прошлая то добавить
                curveBuffers(KeysQ1).size = CalculateSize(.RPaths)
                curveBuffers(KeysQ1).bPut_Long 0
                kof = Int(.MaxLenght * 100000000000#)
                curveBuffers(KeysQ1).bPut_Long kof
                curveBuffers(KeysQ1).bPut_Long UBound(.RPaths)
                For q2 = 1 To UBound(.RPaths)
                    curveBuffers(KeysQ1).bPut_Long UBound(.RPaths(q2).RSegments)
                    For q3 = 1 To UBound(.RPaths(q2).RSegments)
                        With .RPaths(q2).RSegments(q3)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SNPx)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SNPy)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SCPx)
                            curveBuffers(KeysQ1).bPut_Long CLng(.SCPy)
                            curveBuffers(KeysQ1).bPut_Long CLng(.ECPx)
                            curveBuffers(KeysQ1).bPut_Long CLng(.ECPy)
                        End With
                    Next
                Next
            End If
        End With
    Next
    'OKON =============================
    Dim fullBufferSize As Long
    fullBufferSize = keysBuff.size
    For q1 = 1 To UBound(RCurves)
        fullBufferSize = fullBufferSize + curveBuffers(q1).size
    Next
    
    Dim fullBuffer As New buffer
    fullBuffer.size = fullBufferSize + 4
    fullBuffer.InterPlatformNetConverter = True
    fullBuffer.bPut_Long Commands.CurveData
    fullBuffer.bPut_Bytes keysBuff.bytes
    For q1 = 1 To UBound(RCurves)
        fullBuffer.bPut_Bytes curveBuffers(q1).bytes
    Next
    PrepareToSend = fullBuffer.bytes
Err:
End Function

' -----------------------------------------------------------------------------
' CalculateSize (Function), source lines 1638-1644
' -----------------------------------------------------------------------------
Private Function CalculateSize(ByRef RPaths() As RPath) As Long
    CalculateSize = 12
    Dim q1 As Long
    For q1 = 1 To UBound(RPaths)
        CalculateSize = CalculateSize + 4 + UBound(RPaths(q1).RSegments) * 24
    Next
End Function

' -----------------------------------------------------------------------------
' ToStr16bit (Function), source lines 1645-1650
' -----------------------------------------------------------------------------
Private Function ToStr16bit(ByVal n As Long) As String
    Dim aa As Long, bb As Long
    aa = n \ 256
    bb = n - aa * 256
    ToStr16bit = Chr(aa) & Chr(bb)
End Function

' -----------------------------------------------------------------------------
' Odinacovo (Function), source lines 1651-1670
' -----------------------------------------------------------------------------
Private Function Odinacovo(ByRef RCurves1 As RCurve, ByRef RCurves2 As RCurve) As Boolean
    'Проверяет одинаковый ли объект с RCurves1 и RCurves2
    Dim qq As Long, qq2 As Long
    If UBound(RCurves2.RPaths) <> UBound(RCurves1.RPaths) Then Exit Function
    If RCurves1.MaxLenght <> RCurves2.MaxLenght Then Exit Function
    
    For qq = 1 To UBound(RCurves1.RPaths)
    
        If UBound(RCurves2.RPaths(qq).RSegments) <> UBound(RCurves1.RPaths(qq).RSegments) Then Exit Function
        For qq2 = 1 To UBound(RCurves1.RPaths(qq).RSegments)
            If Int(RCurves1.RPaths(qq).RSegments(qq2).ECPx) <> Int(RCurves2.RPaths(qq).RSegments(qq2).ECPx) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).ECPy) <> Int(RCurves2.RPaths(qq).RSegments(qq2).ECPy) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SCPx) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SCPx) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SCPy) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SCPy) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SNPx) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SNPx) Then Exit Function
            If Int(RCurves1.RPaths(qq).RSegments(qq2).SNPy) <> Int(RCurves2.RPaths(qq).RSegments(qq2).SNPy) Then Exit Function
        Next
    Next
    Odinacovo = True
End Function

' -----------------------------------------------------------------------------
' SupRound (Function), source lines 1671-1674
' -----------------------------------------------------------------------------
Private Function SupRound(ByVal Number As Double) As Double
    'Убирает 4 знак после запятой
    SupRound = Int(Number * 10000) / 10000
End Function

' -----------------------------------------------------------------------------
' Combo1_Click (Sub), source lines 1678-1749
' -----------------------------------------------------------------------------
Private Sub Combo1_Click()
    If Combo1DontClick = True Then Exit Sub

    'proverit Change
    If NeedSave = True Then
        If Combo1.ListIndex = Combo1LastSelectListIndex Then
            Combo1DontClick = True
            Combo1.text = Combo1_Text
            Combo1DontClick = False
            Exit Sub
        End If
    
        Dim selectedAnswer As Variant
        selectedAnswer = MsgBox(localization.ActiveLng.Item("RaskroyForm.QSaveChanges") & " (" & Combo1_Text & ")?", vbYesNoCancel + vbQuestion, "PlotCalc")
        If selectedAnswer = vbYes Then
            SaveProfile Combo1LastSelectListIndex
        ElseIf selectedAnswer = vbCancel Then
            Combo1DontClick = True
            Combo1.text = Combo1_Text
            Combo1DontClick = False
            Exit Sub
        End If
    End If
    
    
    Combo1_Text = Combo1.value
    Combo1LastSelectListIndex = Combo1.ListIndex
    CommandSave.Enabled = False
    NeedSave = False
    
    TextFieldsIsEditing = True
    With UserData.RaskroyProfiles(Combo1.ListIndex)
        Text4.text = .MaterialWith / 1000
        Text5.text = .MaterialHeight / 1000
        Check2.value = .Vnutri_Figuri
        Text8.text = .PlotterSheetWith / 1000
        Text7.text = .PlotterSheetHeight / 1000
        Text12.text = .UpField / 1000
        Text13.text = .LeftField / 1000
        Text14.text = .RightField / 1000
        Text15.text = .DownField / 1000
        Check1.value = .Vlogen
        If .GradusCount = 1 Then
            Combo2.ListIndex = 0
        ElseIf .GradusCount = 2 Then
            Combo2.ListIndex = 1
        ElseIf .GradusCount = 4 Then
            Combo2.ListIndex = 2
        ElseIf .GradusCount = 8 Then
            Combo2.ListIndex = 3
        Else
            Combo2.ListIndex = 4
        End If
        Text18.text = .Otstup / 1000
        
        Check4.value = .DrawMaterial
        Check3.value = .DrawWorkObl
        Check5.value = .DrawFields
        Check6.value = .DrawText
        Check7.value = .PlotCalcLayer
        Check8.value = .Group
        Check9.value = .LockObj
    End With
    TextFieldsIsEditing = False
    
    If MultiPage1.SelectedItem.name = "Page2" Then
        DrawRazm ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    Else
        Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    End If
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Combo1_KeyUp (Sub), source lines 1750-1756
' -----------------------------------------------------------------------------
Private Sub Combo1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Combo1_Text <> Combo1.value Then
        CommandSave.Enabled = True
        NeedSave = True
        Combo1_Text = Combo1.value
    End If
End Sub

' -----------------------------------------------------------------------------
' Combo2_Change (Sub), source lines 1757-1762
' -----------------------------------------------------------------------------
Private Sub Combo2_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Command5_Click (Sub), source lines 1764-1766
' -----------------------------------------------------------------------------
Private Sub Command5_Click()
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' CommandNew_Click (Sub), source lines 1768-1816
' -----------------------------------------------------------------------------
Private Sub CommandNew_Click()
    'New profile Button
    
    'proverit Change
    If NeedSave = True Then
        Dim selectedAnswer As Variant
        selectedAnswer = MsgBox(localization.ActiveLng.Item("RaskroyForm.QSaveChanges") & " (" & Combo1_Text & ")?", vbYesNoCancel + vbQuestion, "PlotCalc")
        If selectedAnswer = vbYes Then
            SaveProfile Combo1LastSelectListIndex
        ElseIf selectedAnswer = vbCancel Then
            Combo1DontClick = True
            Combo1.text = Combo1_Text
            Combo1DontClick = False
            Exit Sub
        End If
    End If
    NeedSave = False
    
    
    ReDim Preserve UserData.RaskroyProfiles(UserData.CountRaskroyProfiles)
    UserData.CountRaskroyProfiles = UserData.CountRaskroyProfiles + 1

    Dim newName As String
    newName = CreateNewNameForProfile
    UserData.RaskroyProfiles(UserData.CountRaskroyProfiles - 1).name = newName
    Combo1.AddItem newName, UserData.CountRaskroyProfiles - 1

    With UserData.RaskroyProfiles(UserData.CountRaskroyProfiles - 1)
        .DrawMaterial = True
        .DrawWorkObl = True
        .DrawFields = True
        .DrawText = True
        .Group = True
    End With
    

    Combo1.ListIndex = UserData.CountRaskroyProfiles - 1
    Combo1.Enabled = True
    CommandDelete.Enabled = True
    EnebleProfileAllElements
    UserData.Save
    
    Combo1.SetFocus
    Combo1.SelStart = 0
    Combo1.SelLength = 50
    
    NotNest = True
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' CreateNewNameForProfile (Function), source lines 1817-1840
' -----------------------------------------------------------------------------
Private Function CreateNewNameForProfile(Optional ByVal n As Long = 1) As String
    Dim q As Long
    For q = 0 To UserData.CountRaskroyProfiles - 1
        If n = 1 Then
            If Trim(UserData.RaskroyProfiles(q).name) = localization.ActiveLng.Item("RaskroyForm.NewProfile") Then
                n = n + 1
                CreateNewNameForProfile = CreateNewNameForProfile(n)
                Exit Function
            End If
        Else
            If Trim(UserData.RaskroyProfiles(q).name) = localization.ActiveLng.Item("RaskroyForm.NewProfile") & " " & n Then
                n = n + 1
                CreateNewNameForProfile = CreateNewNameForProfile(n)
                Exit Function
            End If
        End If
    Next
    
    If n = 1 Then
        CreateNewNameForProfile = localization.ActiveLng.Item("RaskroyForm.NewProfile")
    Else
        CreateNewNameForProfile = localization.ActiveLng.Item("RaskroyForm.NewProfile") & " " & n
    End If
End Function

' -----------------------------------------------------------------------------
' CommandSave_Click (Sub), source lines 1842-1849
' -----------------------------------------------------------------------------
Private Sub CommandSave_Click()
    'Save profile Button
    SaveProfile Combo1LastSelectListIndex
    CommandSave.Enabled = False
    NeedSave = False

    'Combo1_Click
End Sub

' -----------------------------------------------------------------------------
' SaveProfile (Sub), source lines 1850-1890
' -----------------------------------------------------------------------------
Private Sub SaveProfile(ByVal i As Long)
    With UserData.RaskroyProfiles(i)
        .name = Combo1_Text
    
        .MaterialWith = ConvertTextToLong(Text4.text)
        .MaterialHeight = ConvertTextToLong(Text5.text)
        .PlotterSheetWith = ConvertTextToLong(Text8.text)
        .PlotterSheetHeight = ConvertTextToLong(Text7.text)
        .UpField = ConvertTextToLong(Text12.text)
        .LeftField = ConvertTextToLong(Text13.text)
        .RightField = ConvertTextToLong(Text14.text)
        .DownField = ConvertTextToLong(Text15.text)
        .Otstup = ConvertTextToLong(Text18.text)
        
        .Vnutri_Figuri = Check2.value
        .Vlogen = Check1.value
        
        If Combo2.ListIndex = 0 Then
            .GradusCount = 1
        ElseIf Combo2.ListIndex = 1 Then
            .GradusCount = 2
        ElseIf Combo2.ListIndex = 2 Then
            .GradusCount = 4
        ElseIf Combo2.ListIndex = 3 Then
            .GradusCount = 8
        Else
            .GradusCount = 32
        End If
        
        .DrawMaterial = Check4.value
        .DrawWorkObl = Check3.value
        .DrawFields = Check5.value
        .DrawText = Check6.value
        .PlotCalcLayer = Check7.value
        .Group = Check8.value
        .LockObj = Check9.value
    End With
    Combo1.List(i) = Combo1_Text
    UserData.LastProfileSelectionNumber = Combo1LastSelectListIndex
    UserData.Save
End Sub

' -----------------------------------------------------------------------------
' ConvertTextToLong (Function), source lines 1891-1897
' -----------------------------------------------------------------------------
Private Function ConvertTextToLong(ByRef value As String) As Long
    On Error GoTo 10
    
    ConvertTextToLong = CLng(val(Replace$(value, ",", ".")) * 1000)
    Exit Function
10
End Function

' -----------------------------------------------------------------------------
' CommandDelete_Click (Sub), source lines 1898-1923
' -----------------------------------------------------------------------------
Private Sub CommandDelete_Click()
    'Delete profile Button
    If MsgBox(localization.ActiveLng.Item("RaskroyForm.QDeleteProfile") & " (" & Trim(UserData.RaskroyProfiles(Combo1LastSelectListIndex).name) & ")?", vbYesNo + vbQuestion, "PlotCalc") = vbNo Then Exit Sub
    
    RemoveProfile Combo1LastSelectListIndex
    Combo1.RemoveItem Combo1LastSelectListIndex
    NeedSave = False
    If UserData.CountRaskroyProfiles > 0 Then
        If Combo1LastSelectListIndex > UserData.CountRaskroyProfiles - 1 Then
            Combo1.ListIndex = UserData.CountRaskroyProfiles - 1
        Else
            Combo1.ListIndex = Combo1LastSelectListIndex
        End If
    Else
        NotNest = True
        RefreshCommand1
        Combo1.text = " " & localization.ActiveLng.Item("RaskroyForm.empty")
        Combo1.Enabled = False
        CommandDelete.Enabled = False
        CommandSave.Enabled = False
        DisebleProfileAllElements
        If MultiPage1.SelectedItem.name = "Page2" Then
            Cls True
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' RemoveProfile (Sub), source lines 1924-1936
' -----------------------------------------------------------------------------
Private Sub RemoveProfile(ByVal i As Long)
    UserData.CountRaskroyProfiles = UserData.CountRaskroyProfiles - 1
    Dim q As Long
    For q = i To UserData.CountRaskroyProfiles - 1
        UserData.RaskroyProfiles(q) = UserData.RaskroyProfiles(q + 1)
    Next
    If UserData.CountRaskroyProfiles > 0 Then
        ReDim Preserve UserData.RaskroyProfiles(UserData.CountRaskroyProfiles - 1)
    Else
        Erase UserData.RaskroyProfiles
    End If
    UserData.Save
End Sub

' -----------------------------------------------------------------------------
' MultiPage1_Change (Sub), source lines 1939-1945
' -----------------------------------------------------------------------------
Private Sub MultiPage1_Change()
    If MultiPage1.SelectedItem.name = "Page2" Then
        DrawRazm ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    Else
        Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    End If
End Sub

' -----------------------------------------------------------------------------
' Text12_Change (Sub), source lines 1947-1953
' -----------------------------------------------------------------------------
Private Sub Text12_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text13_Change (Sub), source lines 1955-1961
' -----------------------------------------------------------------------------
Private Sub Text13_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text14_Change (Sub), source lines 1963-1969
' -----------------------------------------------------------------------------
Private Sub Text14_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text15_Change (Sub), source lines 1971-1977
' -----------------------------------------------------------------------------
Private Sub Text15_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text18_Change (Sub), source lines 1979-1984
' -----------------------------------------------------------------------------
Private Sub Text18_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text4_Change (Sub), source lines 1986-1992
' -----------------------------------------------------------------------------
Private Sub Text4_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text5_Change (Sub), source lines 1994-2000
' -----------------------------------------------------------------------------
Private Sub Text5_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text7_Change (Sub), source lines 2002-2008
' -----------------------------------------------------------------------------
Private Sub Text7_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Text8_Change (Sub), source lines 2010-2016
' -----------------------------------------------------------------------------
Private Sub Text8_Change()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Check1_Click (Sub), source lines 2017-2022
' -----------------------------------------------------------------------------
Private Sub Check1_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' Check3_Click (Sub), source lines 2024-2029
' -----------------------------------------------------------------------------
Private Sub Check3_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    DrawRazm ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
End Sub

' -----------------------------------------------------------------------------
' Check4_Click (Sub), source lines 2031-2036
' -----------------------------------------------------------------------------
Private Sub Check4_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    DrawRazm ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
End Sub

' -----------------------------------------------------------------------------
' Check5_Click (Sub), source lines 2038-2043
' -----------------------------------------------------------------------------
Private Sub Check5_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    DrawRazm ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
End Sub

' -----------------------------------------------------------------------------
' Check6_Click (Sub), source lines 2045-2050
' -----------------------------------------------------------------------------
Private Sub Check6_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    DrawRazm ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
End Sub

' -----------------------------------------------------------------------------
' Check7_Click (Sub), source lines 2052-2056
' -----------------------------------------------------------------------------
Private Sub Check7_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
End Sub

' -----------------------------------------------------------------------------
' Check8_Click (Sub), source lines 2058-2062
' -----------------------------------------------------------------------------
Private Sub Check8_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
End Sub

' -----------------------------------------------------------------------------
' Check9_Click (Sub), source lines 2064-2068
' -----------------------------------------------------------------------------
Private Sub Check9_Click()
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
End Sub

' -----------------------------------------------------------------------------
' TerminateGDI (Sub), source lines 2073-2091
' -----------------------------------------------------------------------------
Private Sub TerminateGDI()
    GdipDeletePen mPen
    GdipDeletePen LinePen
    GdipDeletePen MaterialPen
    GdipDeletePen PagesPen
    GdipDeletePen FieldsPen
    
    GdipDeletePen LinesPaperPen
    GdipDeletePen LinesShadowPen
    
    GdipDeleteBrush hBrush
    GdipDeleteBrush PaperBrush
    GdipDeleteBrush ShadowBrush
    
    
    GdipDeleteGraphics HGraphics
    GdiplusShutdown GdipToken
    ReleaseDC hWndForm, HDCForm
End Sub

' -----------------------------------------------------------------------------
' ConvertColor (Function), source lines 2092-2099
' -----------------------------------------------------------------------------
Private Function ConvertColor(Color As Long, Opacity As Long) As Long
   Dim BGRA(0 To 3) As Byte
   BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
   BGRA(0) = ((Color \ &H10000) And &HFF)
   BGRA(1) = ((Color \ &H100) And &HFF)
   BGRA(2) = (Color And &HFF)
   CopyMemory ConvertColor, BGRA(0), 4&
End Function

' -----------------------------------------------------------------------------
' Check2_Click (Sub), source lines 2100-2154
' -----------------------------------------------------------------------------
Private Sub Check2_Click()
    If Check2.value = False Then
        Frame3.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Label5.Enabled = True
        Label6.Enabled = True
        Label7.Enabled = True
        Text7.Enabled = True
        Text8.Enabled = True
        Text12.Enabled = True
        Text13.Enabled = True
        Text14.Enabled = True
        Text15.Enabled = True
        
        Check4.Enabled = True
        Check3.Enabled = True
        Check5.Enabled = True
        Check6.Enabled = True
        Check7.Enabled = True
        Check8.Enabled = True
        Check9.Enabled = True
    Else
        Frame3.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Label5.Enabled = False
        Label6.Enabled = False
        Label7.Enabled = False
        Text7.Enabled = False
        Text8.Enabled = False
        Text12.Enabled = False
        Text13.Enabled = False
        Text14.Enabled = False
        Text15.Enabled = False
        
        Check4.Enabled = False
        Check3.Enabled = False
        Check5.Enabled = False
        Check6.Enabled = False
        Check7.Enabled = False
        Check8.Enabled = False
        Check9.Enabled = False
    End If
    
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub

' -----------------------------------------------------------------------------
' DisebleAllButtons (Sub), source lines 2155-2164
' -----------------------------------------------------------------------------
Private Sub DisebleAllButtons()
    DisebleProfileAllElements
    Combo1.Enabled = False
    CommandNew.Enabled = False
    CommandSave.Enabled = False
    CommandDelete.Enabled = False
    Command3.Enabled = False
    Command6.Enabled = False
    Frame2.Enabled = False
End Sub

' -----------------------------------------------------------------------------
' EnableAllButtons (Sub), source lines 2165-2175
' -----------------------------------------------------------------------------
Private Sub EnableAllButtons()
    EnebleProfileAllElements
    Combo1.Enabled = True
    CommandNew.Enabled = True
    If NeedSave Then CommandSave.Enabled = True
    CommandDelete.Enabled = True
    If SelectedResult > 1 Then Command3.Enabled = True
    If SelectedResult < CountResults Then Command6.Enabled = True
    If CountResults > 0 And Plotf = False Then Command4.Enabled = True
    Frame2.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' EnableCommand1 (Sub), source lines 2176-2186
' -----------------------------------------------------------------------------
Private Sub EnableCommand1()
    If Connected Then
        If Not Not RCurves Then
            If NotNest Then
                Command1.Enabled = False
            Else
                Command1.Enabled = True
            End If
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' RefreshCommand1 (Sub), source lines 2187-2195
' -----------------------------------------------------------------------------
Private Sub RefreshCommand1()
    Command1.Caption = localization.ActiveLng.Item("RaskroyForm.Command1.Caption")
    Label10.Caption = ""
    Command1.BackColor = VBA.SystemColorConstants.vbButtonFace
    ProgressBar1.GreenColorOn = False
    ImproveOn = False
    EnableCommand1
    Label1.Caption = ".."
End Sub

' -----------------------------------------------------------------------------
' EnebleProfileAllElements (Sub), source lines 2196-2228
' -----------------------------------------------------------------------------
Private Sub EnebleProfileAllElements()
    Frame1.Enabled = True
    If Check2.value = False Then
        Frame3.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Label5.Enabled = True
        Label6.Enabled = True
        Label7.Enabled = True
        Text7.Enabled = True
        Text8.Enabled = True
        Text12.Enabled = True
        Text13.Enabled = True
        Text14.Enabled = True
        Text15.Enabled = True
    End If
    MultiPage1.Enabled = True
    Check2.Enabled = True
    Check1.Enabled = True
    Label8.Enabled = True
    Label9.Enabled = True
    Combo2.Enabled = True
    Text18.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    Check5.Enabled = True
    Check6.Enabled = True
    Check7.Enabled = True
    Check8.Enabled = True
    Check9.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' DisebleProfileAllElements (Sub), source lines 2229-2259
' -----------------------------------------------------------------------------
Private Sub DisebleProfileAllElements()
    MultiPage1.Enabled = False
    Frame1.Enabled = False
    Frame3.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Check2.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    Label7.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Text12.Enabled = False
    Text13.Enabled = False
    Text14.Enabled = False
    Text15.Enabled = False
    Check1.Enabled = False
    Label8.Enabled = False
    Label9.Enabled = False
    Combo2.Enabled = False
    Text18.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Check5.Enabled = False
    Check6.Enabled = False
    Check7.Enabled = False
    Check8.Enabled = False
    Check9.Enabled = False
End Sub

' -----------------------------------------------------------------------------
' ValP (Function), source lines 2260-2262
' -----------------------------------------------------------------------------
Private Function ValP(ByRef txt As String) As Double
    ValP = Abs(val(Replace$(txt, ",", ".")))
End Function

' -----------------------------------------------------------------------------
' Draw (Sub), source lines 2263-2506
' -----------------------------------------------------------------------------
Private Sub Draw(ByVal matW As Double, ByVal matH As Double, _
                 ByVal plotW As Double, ByVal plotH As Double, _
                 ByVal uF As Double, ByVal lF As Double, _
                 ByVal rF As Double, ByVal dF As Double)
    On Error GoTo 50
    NotNest = False
    If CountResults > 0 Then Command3.Enabled = True
    Command6.Enabled = False
    SelectedResult = CountResults + 1
    Command4.Enabled = False
    Cls True
    If Check2.value Then
        GoTo FigurniyList
        Exit Sub
    End If
    
    'Рисует страницу только для ознакомления
    Dim wpix As Double
    Dim hpix As Double
    Dim ks As Double
    Dim dx As Double
    Dim dy As Double
    Dim oUF As Double, oLF As Double
    Dim oRF As Double, oDF As Double
    
    If plotW = 0 Or plotH = 0 Then NotNest = True
    If matW = 0 Or matH = 0 Then
        NotNest = True
        Exit Sub
    End If
    
    
    If matW / WindowWidth > matH / WindowHeigth Then
        ks = (WindowWidth - 15) / matW
    Else
        ks = (WindowHeigth - 20) / matH
    End If
    
    wpix = ks * matW
    hpix = ks * matH
    
    dx = 6
    dy = 8

    FillRect ShadowBrush, dx + 2, WindowHeigth - hpix - dy + 2, wpix, hpix
    FillRect PaperBrush, dx, WindowHeigth - hpix - dy, wpix, hpix
    
    Dim fieldError As Boolean
    Dim str As String
    If lF + rF > plotW Or lF + rF > matW Then
        fieldError = True
    End If
    If uF + dF > plotH Or uF + dF > matH Then
        fieldError = True
    End If
    
    wpix = ks * plotW
    hpix = ks * plotH
    oUF = ks * uF
    oLF = ks * lF
    oRF = ks * rF
    oDF = ks * dF
    If plotW > 0 And plotH > 0 And fieldError = False Then
        If matH / plotH < 100 Then
            Dim ch As Double, c As Double
            For ch = 1 To matH / plotH
                If matW / plotW < 100 Then
                    For c = 1 To matW / plotW
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    Next
                    If matW / plotW > Int(matW / plotW) Then
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    End If
                End If
            Next
            If matH / plotH > Int(matH / plotH) Then
                If matW / plotW < 100 Then
                    For c = 1 To matW / plotW
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    Next
                    If matW / plotW > Int(matW / plotW) Then
                        DrawRect PagesPen, (c - 1) * wpix + dx, -(ch - 0) * hpix + WindowHeigth - dy, wpix, hpix
                        DrawRect FieldsPen, (c - 1) * wpix + dx + oLF, -(ch - 0) * hpix + WindowHeigth - dy + oUF, wpix - oLF - oRF, hpix - oUF - oDF
                    End If
                End If
            End If
        End If
    End If

    wpix = ks * matW
    hpix = ks * matH
    DrawRect LinePen, dx, WindowHeigth - hpix - dy, wpix, hpix
    
    If lF + rF > plotW Or lF + rF > matW Then
        NotNest = True
        str = localization.ActiveLng.Item("RaskroyForm.TooMuchLeftOrRightField")
        TextOut HDCForm, shiftX + 10, shiftY + WindowHeigth - 30, str, Len(str)
    End If
    If uF + dF > plotH Or uF + dF > matH Then
        NotNest = True
        str = localization.ActiveLng.Item("RaskroyForm.TooHighUpperOrLowerField")
        TextOut HDCForm, shiftX + 10, shiftY + WindowHeigth - 30, str, Len(str)
    End If
    
    Exit Sub
FigurniyList:
    '=========================================================
    '           Здесь обработка для фигурного листа
    '=========================================================

    If AlfaWidth / WindowWidth > AlfaHeight / WindowHeigth Then
        ks = (WindowWidth - 15) / AlfaWidth
    Else
        ks = (WindowHeigth - 20) / AlfaHeight
    End If
    
    'NumberOfMaxFigure
    Dim q As Long, q2 As Long
    Dim sNX As Double, sNY As Double
    Dim eNX As Double, eNY As Double
    Dim SCPx As Double, SCPy As Double
    Dim ECPx As Double, ECPy As Double
    Dim minY As Double, maxY As Double
    Dim minX As Double, maxX As Double
    Dim aa As Long
    Dim sheetCurve As PLCurve
    Dim maxRecurse As Long
    Dim a As Double, a1 As Double, a2 As Double
    Dim da1 As Double, da2 As Double
    minX = 999999999999#
    minY = 999999999999#
    
    sheetCurve.CountPaths = UBound(RCurves(NumberOfMaxFigure).RPaths)
    ReDim sheetCurve.Paths(1 To sheetCurve.CountPaths)
    With RCurves(NumberOfMaxFigure)
        For q = 1 To UBound(.RPaths)
            aa = 1
            sNX = .RPaths(q).RSegments(1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
            sNY = .RPaths(q).RSegments(1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                ReDim sheetCurve.Paths(q).segment(1 To 1)
                sheetCurve.Paths(q).segment(1).Xpoint = sNX
                sheetCurve.Paths(q).segment(1).Ypoint = sNY
            With .RPaths(q)
                For q2 = 1 To UBound(.RSegments)
                    SCPx = .RSegments(q2).SCPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    SCPy = .RSegments(q2).SCPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPx = .RSegments(q2).ECPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    ECPy = .RSegments(q2).ECPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    If q2 = UBound(.RSegments) Then
                        eNX = sheetCurve.Paths(q).segment(1).Xpoint
                        eNY = sheetCurve.Paths(q).segment(1).Ypoint
                    Else
                        eNX = .RSegments(q2 + 1).SNPx * RCurves(NumberOfMaxFigure).MaxLenght * ks
                        eNY = .RSegments(q2 + 1).SNPy * RCurves(NumberOfMaxFigure).MaxLenght * ks
                    End If
                        a1 = Angel(sNX, sNY, SCPx, SCPy)
                        a = Angel(SCPx, SCPy, ECPx, ECPy)
                        a2 = Angel(ECPx, ECPy, eNX, eNY)
                        da1 = Abs(DeltaAngle(a, a1))
                        da2 = Abs(DeltaAngle(a2, a))
                        If SCPx = sNX And SCPy = sNY Then da1 = 0
                        If ECPx = eNX And ECPy = eNY Then da2 = 0
                        a = da1 + da2
                        If a < 0.01 Then
                            maxRecurse = 0
                        ElseIf a < 0.05 Then
                            maxRecurse = 1
                        ElseIf a < 0.26 Then
                            maxRecurse = 2
                        ElseIf a < 0.9 Then
                            maxRecurse = 3
                        ElseIf a < 2.2 Then
                            maxRecurse = 4
                        Else
                            maxRecurse = 5
                        End If
                        recursive_bezier sNX, sNY, SCPx, SCPy, ECPx, ECPy, eNX, eNY, 1, sheetCurve.Paths(q), aa, maxRecurse, maxX, minX, maxY, minY
                    sNX = eNX
                    sNY = eNY
                Next
            End With
        Next
    End With
    'Теперь нарисовать лист
    dx = 6
    dy = 8
    For q = 1 To UBound(sheetCurve.Paths)
        With sheetCurve.Paths(q)
            For q2 = 1 To UBound(.segment)
                .segment(q2).Ypoint = -dy + WindowHeigth - .segment(q2).Ypoint + minY
                .segment(q2).Xpoint = dx + .segment(q2).Xpoint - minX
            Next
        End With
    Next
    Dim lines() As LinesString
    Dim y As Long
    Dim y1 As Double, y2 As Double
    Dim x1 As Double, x2 As Double
    ReDim lines(1 To WindowHeigth) As LinesString
    For y = WindowHeigth - (maxY - minY) - dy To WindowHeigth
        For q = 1 To UBound(sheetCurve.Paths)
            With sheetCurve.Paths(q)
                For q2 = 1 To UBound(.segment) - 1
                    y1 = .segment(q2).Ypoint
                    y2 = .segment(q2 + 1).Ypoint
                    If y > y1 And y > y2 Then GoTo nx
                    If y < y1 And y < y2 Then GoTo nx
                    If y2 = y1 Then GoTo nx
                    x1 = .segment(q2).Xpoint
                    x2 = .segment(q2 + 1).Xpoint
                    lines(y).CountX = lines(y).CountX + 1
                    ReDim Preserve lines(y).x(1 To lines(y).CountX)
                    lines(y).x(lines(y).CountX) = (x2 - x1) * (y - y1) / (y2 - y1) + x1
nx:
                Next
            End With
        Next
    Next
    On Error Resume Next
    For q = 1 To UBound(lines)
        With lines(q)
            SortData .x()
        End With
    Next
    For q = 1 To UBound(lines)
        With lines(q)
            For q2 = 1 To .CountX Step 2
                 DrawLine LinesPaperPen, .x(q2), q, .x(q2 + 1), q
                 DrawLine LinesShadowPen, .x(q2) + 2, q + 2, .x(q2 + 1) + 2, q + 2
            Next
        End With
    Next
    For q = 1 To UBound(sheetCurve.Paths)
        With sheetCurve.Paths(q)
            For q2 = 1 To UBound(.segment) - 1
                DrawLine LinePen, .segment(q2).Xpoint, .segment(q2).Ypoint, .segment(q2 + 1).Xpoint, .segment(q2 + 1).Ypoint
            Next
        End With
    Next
50
End Sub

' -----------------------------------------------------------------------------
' Angel (Function), source lines 2507-2515
' -----------------------------------------------------------------------------
Private Function Angel(ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Double
    If x1 = x2 Then
        If y2 > y1 Then Angel = 3.141593
        Exit Function
    End If
    Angel = (y2 - y1) / (x2 - x1)
    Angel = Atn(Angel) + 1.570796
    If (x2 - x1) < 0 Then Angel = Angel + 3.141593
End Function

' -----------------------------------------------------------------------------
' DeltaAngle (Function), source lines 2516-2523
' -----------------------------------------------------------------------------
Private Function DeltaAngle(ByRef a1 As Double, ByRef a2 As Double) As Double
    DeltaAngle = a2 - a1
    If DeltaAngle > 3.141593 Then
        DeltaAngle = DeltaAngle - 6.283186
    ElseIf DeltaAngle < -3.141593 Then
        DeltaAngle = DeltaAngle + 6.283186
    End If
End Function

' -----------------------------------------------------------------------------
' recursive_bezier (Sub), source lines 2524-2563
' -----------------------------------------------------------------------------
Private Sub recursive_bezier(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, R As Double, ByRef path As path, ByRef aa As Long, ByRef maxRecurse As Long, _
                            ByRef maxX As Double, ByRef minX As Double, ByRef maxY As Double, ByRef minY As Double)
   ' Вычислить все средние точек отрезков
    Dim x12 As Double, y12 As Double
    Dim x23 As Double, y23 As Double
    Dim x34 As Double, y34 As Double
    Dim x123 As Double, y123 As Double
    Dim x234 As Double, y234 As Double
    Dim x1234 As Double, y1234 As Double
    
    x12 = (x1 + x2) / 2
    y12 = (y1 + y2) / 2
    x23 = (x2 + x3) / 2
    y23 = (y2 + y3) / 2
    x34 = (x3 + x4) / 2
    y34 = (y3 + y4) / 2
    x123 = (x12 + x23) / 2
    y123 = (y12 + y23) / 2
    x234 = (x23 + x34) / 2
    y234 = (y23 + y34) / 2
    x1234 = (x123 + x234) / 2
    y1234 = (y123 + y234) / 2
    If R > maxRecurse Then
        ' Нарисовать отрезок и закончить
        aa = aa + 1
        ReDim Preserve path.segment(1 To aa)
        path.segment(aa).Xpoint = x4
        path.segment(aa).Ypoint = y4
        If x4 > maxX Then maxX = x4
        If x4 < minX Then minX = x4
        If y4 > maxY Then maxY = y4
        If y4 < minY Then minY = y4
    Else
        ' Продолжить деление
        R = R + 1
        recursive_bezier x1, y1, x12, y12, x123, y123, x1234, y1234, R, path, aa, maxRecurse, maxX, minX, maxY, minY
        recursive_bezier x1234, y1234, x234, y234, x34, y34, x4, y4, R, path, aa, maxRecurse, maxX, minX, maxY, minY
        R = R - 1
   End If
End Sub

' -----------------------------------------------------------------------------
' SortData (Sub), source lines 2564-2581
' -----------------------------------------------------------------------------
Private Sub SortData(ByRef x() As Double)
    If Not Not x() Then
        Dim i As Long
        Dim fSwitched As Boolean
        Dim xTemp As Double
        Do
            fSwitched = False
            For i = 1 To UBound(x) - 1
                If x(i) < x(i + 1) Then 'сортируем по возрастанию
                    fSwitched = True
                    xTemp = x(i)
                    x(i) = x(i + 1)
                    x(i + 1) = xTemp
                End If
            Next
        Loop While fSwitched
    End If
End Sub

' -----------------------------------------------------------------------------
' DrawRazm (Sub), source lines 2582-2647
' -----------------------------------------------------------------------------
Private Sub DrawRazm(ByVal matW As Double, ByVal matH As Double, _
                 ByVal plotW As Double, ByVal plotH As Double, _
                 ByVal uF As Double, ByVal lF As Double, _
                 ByVal rF As Double, ByVal dF As Double)
    On Error Resume Next
    If CountResults > 0 Then Command3.Enabled = True
    Command6.Enabled = False
    SelectedResult = CountResults + 1
    Command4.Enabled = False
    
    If Check2.value Then Cls True: Exit Sub
    
    'Рисует разметку только для ознакомления
    Dim wpix As Double
    Dim hpix As Double
    Dim ks As Double
    Dim dx As Double
    Dim dy As Double
    Dim oUF As Double, oLF As Double
    Dim oRF As Double, oDF As Double
    
    If matW = 0 Or matH = 0 Then Cls True: Exit Sub
    
    FillRect PaperBrush, 0, 0, WindowWidth, WindowHeigth
    DrawRect mPen, 0, 0, WindowWidth, WindowHeigth
    
    If matW / WindowWidth > matH / WindowHeigth Then
        ks = (WindowWidth - 15) / matW
    Else
        ks = (WindowHeigth - 20) / matH
    End If
    
    dx = 6
    dy = 8
    wpix = ks * plotW
    hpix = ks * plotH
    oUF = ks * uF
    oLF = ks * lF
    oRF = ks * rF
    oDF = ks * dF
    
    If wpix > ks * matW Then wpix = ks * matW
    If hpix > ks * matH Then hpix = ks * matH
    Dim wpix2 As Double
    wpix2 = wpix
    If Check3.value Then
        DrawRect PagesPen, dx, WindowHeigth - dy - hpix, wpix, hpix
    End If
    If Check5.value Then
        'Fields
        DrawRect FieldsPen, dx + oLF, WindowHeigth - dy - hpix + oUF, wpix - oRF - oLF, hpix - oUF - oDF
    End If
    
    wpix = ks * matW
    hpix = ks * matH
    If Check4.value Then
        DrawRect MaterialPen, dx, WindowHeigth - dy - hpix, wpix, hpix
    End If
    
    If Check6.value Then
        'Text
        Dim str As String
        str = Round(wpix2 / ks, 0)
        TextOut HDCForm, shiftX + wpix + dx + 2, shiftY + WindowHeigth - dy - 10, str, Len(str)
    End If
End Sub

' -----------------------------------------------------------------------------
' DrawLine (Sub), source lines 2648-2650
' -----------------------------------------------------------------------------
Private Sub DrawLine(ByVal mPen As LongPtr, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    GdipDrawLineI HGraphics, mPen, shiftX + x1, shiftY + y1, shiftX + x2, shiftY + y2
End Sub

' -----------------------------------------------------------------------------
' DrawRect (Sub), source lines 2651-2653
' -----------------------------------------------------------------------------
Private Sub DrawRect(ByVal mPen As LongPtr, ByVal mx As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long)
    GdipDrawRectangleI HGraphics, mPen, shiftX + mx, shiftY + mY, mWidth, mHeight
End Sub

' -----------------------------------------------------------------------------
' FillRect (Sub), source lines 2654-2656
' -----------------------------------------------------------------------------
Private Sub FillRect(ByVal hBrush As LongPtr, ByVal mx As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long)
    GdipFillRectangleI HGraphics, hBrush, shiftX + mx, shiftY + mY, mWidth, mHeight
End Sub

' -----------------------------------------------------------------------------
' Cls (Sub), source lines 2657-2661
' -----------------------------------------------------------------------------
Private Sub Cls(Optional NeedRepaint As Boolean)
    If NeedRepaint Then Me.Repaint
    GdipFillRectangleI HGraphics, hBrush, shiftX, shiftY, WindowWidth, WindowHeigth
    GdipDrawRectangleI HGraphics, mPen, shiftX, shiftY, WindowWidth, WindowHeigth
End Sub

' -----------------------------------------------------------------------------
' Command3_Click (Sub), source lines 2662-2672
' -----------------------------------------------------------------------------
Private Sub Command3_Click()
    SelectedResult = SelectedResult - 1
    If SelectedResult = 1 Then Command3.Enabled = False
    If CountResults > 1 Then Command6.Enabled = True
    If SelectedResult = CountResults Then Command6.Enabled = False
    If Not Plotf Then Command4.Enabled = True
    Label1.Caption = SelectedResult & " " & localization.ActiveLng.Item("RaskroyForm.From") & " " & CountResults & "   (" & Int(Work_Result(SelectedResult).maximumY * 25.4) & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    Dim wrtemp As Work_Result
    wrtemp = Work_Result(SelectedResult)
    DrawWarkResult wrtemp
End Sub

' -----------------------------------------------------------------------------
' Command6_Click (Sub), source lines 2673-2682
' -----------------------------------------------------------------------------
Private Sub Command6_Click()
    SelectedResult = SelectedResult + 1
    If SelectedResult = CountResults Then Command6.Enabled = False
    Command3.Enabled = True
    If Not Plotf Then Command4.Enabled = True
    Label1.Caption = SelectedResult & " " & localization.ActiveLng.Item("RaskroyForm.From") & " " & CountResults & "   (" & Int(Work_Result(SelectedResult).maximumY * 25.4) & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    Dim wrtemp As Work_Result
    wrtemp = Work_Result(SelectedResult)
    DrawWarkResult wrtemp
End Sub

' -----------------------------------------------------------------------------
' FindPostion (Sub), source lines 2683-2703
' -----------------------------------------------------------------------------
Private Sub FindPostion(ByRef PositionX As Double, ByRef PositionY As Double, SingMaterialWith As Double, SingMaterialHeight As Double, rightGo As Boolean)
    Dim q As Long, chang As Boolean
    Do
        chang = False
        For q = 1 To VGCore.ActiveDocument.ActivePage.shapes.count
            With VGCore.ActiveDocument.ActivePage.shapes(q)
                If .Type <> cdrGuidelineShape Then
                    If (PositionX < .PositionX + .SizeWidth And PositionX + SingMaterialWith > .PositionX) _
                        And (PositionY < .PositionY And PositionY + SingMaterialWith > .PositionY - .SizeHeight) Then
                        chang = True
                        If rightGo = True Then
                            PositionX = .PositionX + .SizeWidth + SingMaterialWith / 20
                        Else
                            PositionY = .PositionY + SingMaterialWith / 20
                        End If
                    End If
                End If
            End With
        Next
    Loop While chang
End Sub

' -----------------------------------------------------------------------------
' IsFreePlaceOnThePage (Function), source lines 2704-2721
' -----------------------------------------------------------------------------
Private Function IsFreePlaceOnThePage(ByRef PositionX As Double, ByRef PositionY As Double, SingMaterialWith As Double, SingMaterialHeight As Double) As Boolean
    Dim q As Long

    For q = 1 To VGCore.ActiveDocument.ActivePage.shapes.count
        With VGCore.ActiveDocument.ActivePage.shapes(q)
            If .Type <> cdrGuidelineShape Then
                If (PositionX < .PositionX + .SizeWidth And PositionX + SingMaterialWith > .PositionX) _
                    And (PositionY < .PositionY And PositionY + SingMaterialWith > .PositionY - .SizeHeight) Then
                    If Not FindStaticIdInDu(.StaticId) Then
                        IsFreePlaceOnThePage = False
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
    IsFreePlaceOnThePage = True
End Function

' -----------------------------------------------------------------------------
' FindStaticIdInDu (Function), source lines 2722-2735
' -----------------------------------------------------------------------------
Private Function FindStaticIdInDu(StaticId As Long) As Boolean
    Dim q As Long
    Dim q2 As Long
    For q = 1 To UBound(Du)
        With Du(q)
            For q2 = 1 To .shapes.count
                If .shapes(q2).StaticId = StaticId Then
                    FindStaticIdInDu = True
                    Exit Function
                End If
            Next
        End With
    Next
End Function

' -----------------------------------------------------------------------------
' Okrugl (Sub), source lines 2736-2818
' -----------------------------------------------------------------------------
Private Sub Okrugl(ByRef Work_Result As Work_Result)
    '===========================================
    'Округляет результаты (размеры страниц)
    '===========================================
    Dim q As Long, q2 As Long, q3 As Long
    With Work_Result
        If .Materials(1).RLines(1).CountRSheets = 1 Then
            For q = 1 To .CountMaterials
                With .Materials(q)
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            .RSheets(1).Width = Work_Result.SingMaterialWith - Work_Result.SingLeftField - Work_Result.SingRightField
                        End With
                    Next
                End With
            Next
        End If
    End With
    'теперь округление
    Dim ostatok  As Double
    Dim d As Double, d2 As Double, d3 As Double
    With Work_Result
            For q = 1 To .CountMaterials
                With .Materials(q)
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            'Находим остаток
                            ostatok = Work_Result.SingMaterialWith
                            For q3 = 1 To .CountRSheets
                                With .RSheets(q3)
                                    ostatok = ostatok - .Width - Work_Result.SingLeftField - Work_Result.SingRightField
                                End With
                            Next
                            'Тепреь распределим остаток на все страницы так чтобы округлить их
                            For q3 = 1 To .CountRSheets
                                With .RSheets(q3)
                                    d = ((1 - ((.Width * 25.4) / 10 - Int((.Width * 25.4) / 10))) * 10) / 25.4  'то сколько нужно прибывить (кратно 10)
                                    d2 = ((1 - ((.Width * 25.4) / 5 - Int((.Width * 25.4) / 5))) * 5) / 25.4    'то сколько нужно прибывить (кратно 5)
                                    d3 = ((1 - ((.Width * 25.4) / 1 - Int((.Width * 25.4) / 1))) * 1) / 25.4    'то сколько нужно прибывить (кратно 1)
                                    If ostatok > d And (.Width + d <= Work_Result.SingPlotterSheetWith) Then
                                        ostatok = ostatok - d
                                        .Width = .Width + d
                                    ElseIf ostatok > d2 And (.Width + d2 <= Work_Result.SingPlotterSheetWith) Then
                                         ostatok = ostatok - d2
                                        .Width = .Width + d2
                                    ElseIf ostatok > d3 And (.Width + d3 <= Work_Result.SingPlotterSheetWith) Then
                                         ostatok = ostatok - d3
                                        .Width = .Width + d3
                                    End If
                                End With
                            Next
                        End With
                    Next
                    'теперь вертикальное округление
                    ostatok = Work_Result.SingMaterialHeight
                    'находим остаток
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            ostatok = ostatok - .Height - Work_Result.SingUpField - Work_Result.SingDownField
                        End With
                    Next
                    
                    For q2 = 1 To .CountRLines
                        With .RLines(q2)
                            d = ((1 - ((.Height * 25.4) / 10 - Int((.Height * 25.4) / 10))) * 10) / 25.4    'то сколько нужно прибывить (кратно 10)
                            d2 = ((1 - ((.Height * 25.4) / 5 - Int((.Height * 25.4) / 5))) * 5) / 25.4      'то сколько нужно прибывить (кратно 5)
                            d3 = ((1 - ((.Height * 25.4) / 1 - Int((.Height * 25.4) / 1))) * 1) / 25.4      'то сколько нужно прибывить (кратно 1)
                            If ostatok > d And (.Height + d <= Work_Result.SingPlotterSheetHeight) Then
                                ostatok = ostatok - d
                                .Height = .Height + d
                            ElseIf ostatok > d2 And (.Height + d2 <= Work_Result.SingPlotterSheetHeight) Then
                                ostatok = ostatok - d2
                                .Height = .Height + d2
                            ElseIf ostatok > d3 And (.Height + d3 <= Work_Result.SingPlotterSheetHeight) Then
                                ostatok = ostatok - d3
                                .Height = .Height + d3
                            End If
                        End With
                    Next
                End With
            Next
    End With
End Sub

' -----------------------------------------------------------------------------
' Razmestit_Reazultat (Sub), source lines 2819-3055
' -----------------------------------------------------------------------------
Private Sub Razmestit_Reazultat(ByRef Work_Result As Work_Result, _
                         DrawMaterial As Boolean, _
                         DrawFields As Boolean, _
                         DrawWorkObl As Boolean, _
                         DrawText As Boolean, _
                         PlotCalcLayer As Boolean, _
                         Group As Boolean, _
                         LockObj As Boolean, _
                         OnPages As Boolean)
    '===================================================
    'Размещает фигуры в CorelDRAW
    '===================================================
    On Error GoTo 20
    VGCore.Optimization = True
    VGCore.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("RaskroyForm.BeginCommandGroup")
    Okrugl Work_Result 'Округление размеров
    Dim pi As Double: pi = 3.14159265
    Dim plusX As Double
    Dim plusY As Double
    Dim lastPX As Double, lastPY As Double
    Dim sheetX As Double, sheetY As Double
    Dim text2 As String
    Dim dx As Double, dy As Double
    Dim rightGo As Boolean
    Dim shiftX As Double, shiftY As Double
    Dim q As Long, m As Long, rl As Long, rs As Long
    
    Dim RazmetkaObject As VGCore.shape
    Dim RazmetkaObjects As New VGCore.ShapeRange
    Dim RaskroyObjects As New VGCore.ShapeRange
    
    Dim RazmLayer As VGCore.Layer
    Dim RazmLayerStr As String

    Dim CurPage As VGCore.page
    
    If OnPages Then
        With VGCore.ActiveDocument.ActivePage
            .SizeWidth = Work_Result.SingMaterialWith
            .SizeHeight = Work_Result.SingMaterialHeight
        
            If Not IsFreePlaceOnThePage(.LeftX, .BottomY, .SizeWidth, .SizeHeight) Then
                'Insert New Page
                ActiveDocument.InsertPagesEx 1, False, .index, Work_Result.SingMaterialWith, Work_Result.SingMaterialHeight
            End If
        End With
    Else
        lastPX = VGCore.ActiveSelection.PositionX + VGCore.ActiveSelection.SizeWidth + Work_Result.SingMaterialWith / 20
        lastPY = VGCore.ActiveSelection.PositionY - VGCore.ActiveSelection.SizeHeight
    End If
    
    If Work_Result.SingMaterialWith > Work_Result.SingMaterialHeight Then
        plusY = Work_Result.SingMaterialHeight + Work_Result.SingMaterialHeight / 20
        rightGo = False
    Else
        plusX = Work_Result.SingMaterialWith + Work_Result.SingMaterialWith / 20
        rightGo = True
    End If
    
    With Work_Result
        For m = 1 To .CountMaterials
            sheetY = 0
            With .Materials(m)
                If OnPages Then
                    .PositionX = VGCore.ActiveDocument.ActivePage.LeftX
                    .PositionY = VGCore.ActiveDocument.ActivePage.BottomY
                Else
                    If m > 1 Then
                        .PositionX = lastPX + plusX
                        .PositionY = lastPY + plusY
                    Else
                        .PositionX = lastPX
                        .PositionY = lastPY
                    End If
                    'Проверить свободна ли позиция
                    FindPostion .PositionX, .PositionY, Work_Result.SingMaterialWith, Work_Result.SingMaterialHeight, rightGo
                    lastPX = .PositionX
                    lastPY = .PositionY
                End If
                
                If DrawMaterial Then
                    Set RazmetkaObject = VGCore.ActiveLayer.CreateRectangle(.PositionX, .PositionY, Work_Result.SingMaterialWith + .PositionX, Work_Result.SingMaterialHeight + .PositionY) 'Это весь материал
                    RazmetkaObject.Outline.SetProperties Color:=CreateRGBColor(102, 102, 204)
                    RazmetkaObjects.Add RazmetkaObject
                End If
                For rl = 1 To .CountRLines
                    sheetX = 0
                    With .RLines(rl)
                        'текст==========================================================================
                        text2 = "/"
                        If DrawText Then
                            If .CountRSheets = 1 Then
                                If .RSheets(1).CountRFigures > 0 Then
                                    Set RazmetkaObject = VGCore.ActiveLayer.CreateArtisticText(0, 0, Int((.Height + Work_Result.SingDownField + Work_Result.SingUpField) * 25.4 + 0.5))
                                    RazmetkaObject.Rotate 90
                                    RazmetkaObject.Stretch Work_Result.SingMaterialHeight / 20
                                    RazmetkaObject.PositionY = sheetY + Work_Result.Materials(m).PositionY + (.Height + Work_Result.SingDownField + Work_Result.SingUpField) / 2 + RazmetkaObject.SizeHeight / 2
                                    RazmetkaObject.PositionX = Work_Result.Materials(m).PositionX - RazmetkaObject.SizeWidth - RazmetkaObject.SizeWidth / 2
                                    RazmetkaObjects.Add RazmetkaObject
                                End If
                            Else
                                If .RSheets(1).CountRFigures > 0 Then
                                    Set RazmetkaObject = VGCore.ActiveLayer.CreateArtisticText(0, 0, Int((.Height + Work_Result.SingDownField + Work_Result.SingUpField) * 25.4 + 0.5))
                                    RazmetkaObject.Stretch Work_Result.SingMaterialHeight / 20
                                    RazmetkaObject.Rotate 270
                                    RazmetkaObject.PositionY = sheetY + Work_Result.Materials(m).PositionY + (.Height + Work_Result.SingDownField + Work_Result.SingUpField) / 2 + RazmetkaObject.SizeHeight / 2
                                    RazmetkaObject.PositionX = Work_Result.Materials(m).PositionX + Work_Result.SingMaterialWith + RazmetkaObject.SizeWidth / 2
                                    dx = RazmetkaObject.PositionX + RazmetkaObject.SizeWidth
                                    dy = RazmetkaObject.PositionY - RazmetkaObject.SizeHeight
                                    RazmetkaObjects.Add RazmetkaObject
                                End If
                            End If
                        End If '========================================================================
                        For rs = 1 To .CountRSheets
                            With .RSheets(rs)
                                If .CountRFigures > 0 Then
                                    text2 = text2 & Int((Work_Result.SingLeftField + .Width + Work_Result.SingRightField) * 25.4 + 0.5) & "/"
                                    If DrawFields Then

                                        Set RazmetkaObject = VGCore.ActiveLayer.CreateRectangle(Work_Result.Materials(m).PositionX + Work_Result.SingLeftField + sheetX, _
                                                                                    Work_Result.Materials(m).PositionY + Work_Result.SingDownField + sheetY, _
                                                                                   .Width + Work_Result.Materials(m).PositionX + Work_Result.SingLeftField + sheetX, _
                                                                                    Work_Result.Materials(m).RLines(rl).Height + Work_Result.Materials(m).PositionY + Work_Result.SingDownField + sheetY) 'Это весь материал
                                        RazmetkaObject.Outline.SetProperties Color:=CreateRGBColor(255, 102, 0)
                                        RazmetkaObject.Outline.SetProperties 0.059055
                                        RazmetkaObjects.Add RazmetkaObject
                                    End If
                                    If DrawWorkObl Then
                                        Set RazmetkaObject = VGCore.ActiveLayer.CreateRectangle(Work_Result.Materials(m).PositionX + sheetX, _
                                                                                        Work_Result.Materials(m).PositionY + sheetY, _
                                                                                       .Width + Work_Result.Materials(m).PositionX + Work_Result.SingLeftField + Work_Result.SingRightField + sheetX, _
                                                                                        Work_Result.Materials(m).RLines(rl).Height + Work_Result.Materials(m).PositionY + Work_Result.SingDownField + Work_Result.SingUpField + sheetY) 'Это весь материал
                                        RazmetkaObject.Outline.SetProperties Color:=CreateRGBColor(255, 102, 0)
                                        RazmetkaObjects.Add RazmetkaObject
                                    End If
                                End If
                                
                                For q = 1 To .CountRFigures
                                    With .RFigures(q)
                                        Rerotator shiftX, shiftY, RerotateCrv(.Q_figure), Du(.Q_figure), (-.gradus / (2 * pi / 360))
                                        Du(.Q_figure).PositionX = .x - Du(.Q_figure).SizeWidth + Work_Result.SingLeftField + Work_Result.Materials(m).PositionX + sheetX - 2 * Work_Result.Otstup - shiftX
                                        Du(.Q_figure).PositionY = .y + Work_Result.SingDownField + Work_Result.Materials(m).PositionY + sheetY - 2 * Work_Result.Otstup - shiftY
                                        RaskroyObjects.AddRange Du(.Q_figure)
                                    End With
                                Next
                                
                                
                                sheetX = sheetX + Work_Result.SingLeftField + .Width + Work_Result.SingRightField
                            End With
                        Next
                        sheetY = sheetY + Work_Result.SingDownField + .Height + Work_Result.SingUpField
                        'текст горизонтальный
                        If DrawText And .CountRSheets > 1 And .RSheets(1).CountRFigures > 0 Then
                            Set RazmetkaObject = ActiveLayer.CreateArtisticText(dx, dy, text2)
                            RazmetkaObject.Stretch Work_Result.SingMaterialHeight / 20
                            RazmetkaObjects.Add RazmetkaObject
                        End If
                        
                        
                        
                    End With
                Next
            End With
            'RazmLayerStr
            If OnPages Then
                If m = 1 Then
                    Set CurPage = VGCore.ActivePage
                    
                    If PlotCalcLayer Then
                        VGCore.ActivePage.CreateLayer "PlotCalc"
                        RazmLayerStr = "PlotCalc"
                        
                    Else
                        RazmLayerStr = VGCore.ActiveLayer.name
                    End If
                Else ' m > 1
                    Set CurPage = ActiveDocument.InsertPagesEx(1, False, VGCore.ActivePage.index, .SingMaterialWith, .SingMaterialHeight)
                    
                End If
                
                
                RazmetkaObjects.MoveToLayer CurPage.Layers(RazmLayerStr)
                For q = 1 To RaskroyObjects.count
                    With RaskroyObjects(q)
                        .MoveToLayer CurPage.Layers(.Layer.name)
                    End With
                Next
                
                VGCore.ActiveDocument.ClearSelection
                If RazmetkaObjects.count > 0 Then
                    RazmetkaObjects.AddToSelection
                    If Group Then
                        VGCore.ActiveSelection.Group
                    End If
                    If LockObj Then
                        VGCore.ActiveSelection.Locked = True
                    End If
                End If
                
                RazmetkaObjects.RemoveAll
                RaskroyObjects.RemoveAll
            End If
        Next
    End With
    
    VGCore.ActiveDocument.ClearSelection
    
    If Not OnPages Then
        If RazmetkaObjects.count > 0 Then
            'Layer
            
            If PlotCalcLayer Then
                Set RazmLayer = VGCore.ActivePage.CreateLayer("PlotCalc")
                RazmetkaObjects.MoveToLayer RazmLayer
            End If
            
            RazmetkaObjects.AddToSelection
            If Group Then
                VGCore.ActiveSelection.Group
            End If
            If LockObj Then
                VGCore.ActiveSelection.Locked = True
            End If
        End If
        
        'Добаить в выделение ВСЁ
        VGCore.ActiveDocument.ClearSelection
        RazmetkaObjects.AddToSelection
        RaskroyObjects.AddToSelection
    End If


20
    VGCore.Optimization = False
    VGCore.ActiveDocument.EndCommandGroup
    VGCore.Refresh
End Sub

' -----------------------------------------------------------------------------
' Rerotator (Sub), source lines 3058-3084
' -----------------------------------------------------------------------------
Private Sub Rerotator(ByRef shiftX As Double, ByRef shiftY As Double, crv As curve, sr As ShapeRange, Angel As Double)
    Dim transformMatrix As New transformMatrix
    transformMatrix.Rotate Angel
      
    Dim mfc() As MatrixForCheck
    mfc = GetMatrixForCheck(sr)
    
    Dim q As Long

    
    For q = 1 To sr.shapes.count
        sr.shapes(q).ApplyTransformMatrix transformMatrix
    Next
    
    Dim mfc2() As MatrixForCheck
    mfc2 = GetMatrixForCheck(sr)
    

    ApplySecondTransformMatrix sr, mfc, transformMatrix
    

    crv.ApplyTransformMatrix transformMatrix
        
    shiftX = crv.boundingBox.right - sr.RightX
    shiftY = crv.boundingBox.top - sr.PositionY
    
End Sub

' -----------------------------------------------------------------------------
' ApplySecondTransformMatrix (Sub), source lines 3085-3103
' -----------------------------------------------------------------------------
Private Sub ApplySecondTransformMatrix(sr As ShapeRange, ByRef mfc() As MatrixForCheck, ByRef transformMatrix As transformMatrix, Optional ByRef qq As Long = -1)
    
    Dim q As Long

    For q = 1 To sr.shapes.count
        qq = qq + 1
        'Compare
        If CompareMatrixes(mfc(qq), sr.shapes(q)) Then
            'Apply
            sr.shapes(q).ApplyTransformMatrix transformMatrix
        End If
        
        If IsPowerClip(sr.shapes(q)) Then
            ApplySecondTransformMatrix sr.shapes(q).PowerClip.shapes.All, mfc(), transformMatrix, qq
        ElseIf sr.shapes(q).Type = cdrGroupShape Then
            ApplySecondTransformMatrix sr.shapes(q).shapes.All, mfc(), transformMatrix, qq
        End If
    Next
End Sub

' -----------------------------------------------------------------------------
' CompareMatrixes (Function), source lines 3104-3112
' -----------------------------------------------------------------------------
Private Function CompareMatrixes(m1 As MatrixForCheck, m2 As VGCore.shape) As Boolean
    With m2.TransformationMatrix
        If m1.d11 <> .d11 Then Exit Function
        If m1.d12 <> .d12 Then Exit Function
        If m1.d21 <> .d21 Then Exit Function
        If m1.d22 <> .d22 Then Exit Function
    End With
    CompareMatrixes = True
End Function

' -----------------------------------------------------------------------------
' GetMatrixForCheck (Function), source lines 3113-3138
' -----------------------------------------------------------------------------
Private Function GetMatrixForCheck(sr As ShapeRange) As MatrixForCheck()
    Dim mfc() As MatrixForCheck
    Dim q As Long

    For q = 1 To sr.shapes.count
        If Not Not mfc() Then
            ReDim Preserve mfc(UBound(mfc()) + 1)
        Else
            ReDim mfc(0)
        End If
        
        With sr.shapes(q).TransformationMatrix
            mfc(UBound(mfc)).d11 = .d11
            mfc(UBound(mfc)).d12 = .d12
            mfc(UBound(mfc)).d21 = .d21
            mfc(UBound(mfc)).d22 = .d22
        End With
        
        If IsPowerClip(sr.shapes(q)) Then
            AddArrayToArray mfc, GetMatrixForCheck(sr.shapes(q).PowerClip.shapes.All)
        ElseIf sr.shapes(q).Type = cdrGroupShape Then
            AddArrayToArray mfc, GetMatrixForCheck(sr.shapes(q).shapes.All)
        End If
    Next
    GetMatrixForCheck = mfc()
End Function

' -----------------------------------------------------------------------------
' AddArrayToArray (Sub), source lines 3139-3157
' -----------------------------------------------------------------------------
Private Sub AddArrayToArray(ByRef arr() As MatrixForCheck, arr2() As MatrixForCheck)
    If Not Not arr() Then
        If Not Not arr2() Then
            Dim arr2count As Long
            Dim q As Long, c As Long
            arr2count = UBound(arr2()) + 1
            q = UBound(arr)
            ReDim Preserve arr(q + arr2count)
            For q = q + 1 To q + arr2count
                arr(q) = arr2(c)
                c = c + 1
            Next
        End If
    Else
        If Not Not arr2() Then
            arr() = arr2()
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' IsPowerClip (Function), source lines 3158-3162
' -----------------------------------------------------------------------------
Private Function IsPowerClip(sh As shape) As Boolean
    On Error GoTo 10
    If sh.PowerClip.shapes.count > 0 Then IsPowerClip = True: Exit Function
10:
End Function

' -----------------------------------------------------------------------------
' Command4_Click (Sub), source lines 3164-3168
' -----------------------------------------------------------------------------
Private Sub Command4_Click()
    '[OK]
    OkPress = True
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' UserForm_QueryClose (Sub), source lines 3172-3205
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 3206-3220
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    'proverit Change
    Break = True
    If UserData.LastProfileSelectionNumber <> Combo1LastSelectListIndex Then
        UserData.LastProfileSelectionNumber = Combo1LastSelectListIndex
        UserData.Save
    End If
    TerminateGDI
    Sleep 200
    BlocksSocket1.Disconnect


    Application.ActiveDocument.Unit = lastUnit
    ActiveDocument.ReferencePoint = lastReferencePoint
End Sub
