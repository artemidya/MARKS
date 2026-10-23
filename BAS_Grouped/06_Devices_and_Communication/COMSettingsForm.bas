' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: COMSettingsForm.frm
' Procedures: 26
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 66-99
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Список распространённых скоростей передачи данных
    Dim baudRates As Variant
    baudRates = Array(110, 300, 1200, 2400, 4800, 9600, _
                      19200, 38400, 57600, 115200, 230400, 460800, 921600)
    
    Dim i As Integer
    For i = LBound(baudRates) To UBound(baudRates)
        Combo1.AddItem baudRates(i)
    Next i

    ComboBox1.AddItem 5
    ComboBox1.AddItem 6
    ComboBox1.AddItem 7
    ComboBox1.AddItem 8
    
    ComboBox2.AddItem "None"
    ComboBox2.AddItem "Odd"
    ComboBox2.AddItem "Even"
    ComboBox2.AddItem "Mark"
    ComboBox2.AddItem "Space"
    
    ComboBox3.AddItem "1"
    ComboBox3.AddItem "1.5"
    ComboBox3.AddItem "2"
    
    ComboBox4.AddItem "None"
    ComboBox4.AddItem "Xon/Xoff"
    ComboBox4.AddItem "RTS/CTS"
    ComboBox4.AddItem "DSR/DTR"
    
    
   
End Sub

' -----------------------------------------------------------------------------
' UserForm_Activate (Sub), source lines 100-152
' -----------------------------------------------------------------------------
Private Sub UserForm_Activate()

    Dim RegData1 As New RegData
    Dim ComSettingsStr As String
    '"baud=9600 parity=N data=8 stop=1"
    ComSettingsStr = RegData1.ComSettingsStr(PortName)
    dntChange = True
    Combo1.ListIndex = GetBaudListIndex(Combo1, ComSettingsStr)
    ComboBox1.ListIndex = GetDataBitsListIndex(ComboBox1, ComSettingsStr)
    ComboBox2.ListIndex = GetParityListIndex(ComboBox2, ComSettingsStr)
    ComboBox3.ListIndex = GetStopBitsListIndex(ComboBox3, ComSettingsStr)
    ComboBox4.ListIndex = GetFlowControlListIndex(ComboBox4, ComSettingsStr)
    
    
    dntChange = False
    
    isActive = True
    'localization.ActiveLng.Item("COMSettingsForm.")
    
    COMSettingsForm.Caption = localization.ActiveLng.Item("COMSettingsForm.Caption") & " " & PortName  '"Свойства:"
    Label1.Caption = PortName
    Frame1.Caption = localization.ActiveLng.Item("COMSettingsForm.Frame1.Caption") '"Параметры порта"
    Label2.Caption = localization.ActiveLng.Item("COMSettingsForm.Label2.Caption") '"Скорость (бит/с):"
    Label3.Caption = localization.ActiveLng.Item("COMSettingsForm.Label3.Caption") '"Биты данных:"
    Label4.Caption = localization.ActiveLng.Item("COMSettingsForm.Label4.Caption") '"Чётность:"
    Label5.Caption = localization.ActiveLng.Item("COMSettingsForm.Label5.Caption") '"Стоповые биты:"
    Label6.Caption = localization.ActiveLng.Item("COMSettingsForm.Label6.Caption") '"Управление потоком:"
    CommandButton3.Caption = localization.ActiveLng.Item("COMSettingsForm.CommandButton3.Caption") '"Дополнительно.."
    Command5.Caption = localization.ActiveLng.Item("COMSettingsForm.Command5.Caption") '"Восстановить значения по умолчанию"
    Frame2.Caption = localization.ActiveLng.Item("COMSettingsForm.Frame2.Caption") '"Проверка порта"
    CommandButton1.Caption = localization.ActiveLng.Item("PlotterForm.CommandButton22.Caption") & " HPGL"
    CommandButton2.Caption = localization.ActiveLng.Item("PlotterForm.CommandButton22.Caption") & " DMPL"
    CommandButton1.ControlTipText = localization.ActiveLng.Item("PlotterForm.CommandButton22.ControlTipText")
    CommandButton2.ControlTipText = localization.ActiveLng.Item("PlotterForm.CommandButton22.ControlTipText")
    
    If AdvancedSettings Then
        CommandButton3.Enabled = True
    Else
        CommandButton3.Enabled = False
    End If
    
    If PlotterForm.OutPuts.Connected(PlotterForm.Combo1.ListIndex) Then
        Frame2.Enabled = True
        CommandButton1.Enabled = True
        CommandButton2.Enabled = True
    Else
        Frame2.Enabled = False
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
    End If
    CheckDefaultButton
    ForEventsLoop
End Sub

' -----------------------------------------------------------------------------
' Combo1_Change (Sub), source lines 154-158
' -----------------------------------------------------------------------------
Private Sub Combo1_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub

' -----------------------------------------------------------------------------
' ComboBox1_Change (Sub), source lines 160-164
' -----------------------------------------------------------------------------
Private Sub ComboBox1_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub

' -----------------------------------------------------------------------------
' ComboBox2_Change (Sub), source lines 166-170
' -----------------------------------------------------------------------------
Private Sub ComboBox2_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub

' -----------------------------------------------------------------------------
' ComboBox3_Change (Sub), source lines 172-176
' -----------------------------------------------------------------------------
Private Sub ComboBox3_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub

' -----------------------------------------------------------------------------
' ComboBox4_Change (Sub), source lines 178-182
' -----------------------------------------------------------------------------
Private Sub ComboBox4_Change()
    If dntChange Then Exit Sub
    BuildComSettingsString
    CheckDefaultButton
End Sub

' -----------------------------------------------------------------------------
' CommandButton1_Click (Sub), source lines 185-192
' -----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    Dim txt As String
    Dim bytes() As Byte
    
    txt = "IN;PA;VS5;PU0,0;PU0,400;PU400,400;PU400,0;PU0,0;"
    bytes = PlotterLanguage.StrToAscii(txt)
    PlotterForm.OutPuts.Out bytes, PlotterForm.OutPuts.x, PlotterForm.OutPuts.y, True
End Sub

' -----------------------------------------------------------------------------
' CommandButton2_Click (Sub), source lines 194-201
' -----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    Dim txt As String
    Dim bytes() As Byte
    
    txt = ";:HL0,EC1,UP1,AV5,U0,0,U0,400,U400,400,U400,0,U0,0,"
    bytes = PlotterLanguage.StrToAscii(txt)
    PlotterForm.OutPuts.Out bytes, PlotterForm.OutPuts.x, PlotterForm.OutPuts.y, True
End Sub

' -----------------------------------------------------------------------------
' ForEventsLoop (Sub), source lines 204-214
' -----------------------------------------------------------------------------
Private Sub ForEventsLoop()
    Do While isActive
        For q = 0 To 6
            Sleep 20
            DoEvents
        Next
        If Not PlotterForm.USB_Events1 Is Nothing Then
            PlotterForm.USB_Events1.USB_change_check_loop
        End If
    Loop
End Sub

' -----------------------------------------------------------------------------
' Command5_Click (Sub), source lines 216-227
' -----------------------------------------------------------------------------
Private Sub Command5_Click()
    dntChange = True
    ' Установка значения по умолчанию
    Combo1.value = 9600
    ComboBox1.value = 8
    ComboBox2.value = "None"
    ComboBox3.value = "1"
    ComboBox4.value = "None"
    dntChange = False
    BuildComSettingsString
    CheckDefaultButton
End Sub

' -----------------------------------------------------------------------------
' CheckDefaultButton (Sub), source lines 228-238
' -----------------------------------------------------------------------------
Private Sub CheckDefaultButton()
    If Combo1.value = 9600 And _
    ComboBox1.value = 8 And _
    ComboBox2.value = "None" And _
    ComboBox3.value = "1" And _
    ComboBox4.value = "None" Then
        Command5.Enabled = False
    Else
        Command5.Enabled = True
    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton3_Click (Sub), source lines 240-261
' -----------------------------------------------------------------------------
Private Sub CommandButton3_Click()
    On Error GoTo 10
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    si.cb = LenB(si)
    Dim result As Long
        
    result = CreateProcessA( _
        vbNullString, _
        "rundll32.exe" & " devmgr.dll,DeviceProperties_RunDLL /DeviceID """ & deviceId & """", _
        0, 0, 0, 0, 0, _
        vbNullString, _
        si, pi)

    If result <> 0 Then
        ' Ожидание завершения внешнего процесса (модальное поведение)
        WaitForSingleObject pi.hProcess, -1
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
    End If
10
End Sub

' -----------------------------------------------------------------------------
' BuildComSettingsString (Sub), source lines 262-272
' -----------------------------------------------------------------------------
Private Sub BuildComSettingsString()
    Dim result As String
    result = Trim(BuildBaudPart(Combo1) & " " & _
             BuildParityPart(ComboBox2) & " " & _
             BuildDataBitsPart(ComboBox1) & " " & _
             BuildStopBitsPart(ComboBox3) & " " & _
             BuildFlowControlParts(ComboBox4))

    Dim RegData1 As New RegData
    RegData1.ComSettingsStr(PortName) = result
End Sub

' -----------------------------------------------------------------------------
' BuildBaudPart (Function), source lines 273-279
' -----------------------------------------------------------------------------
Private Function BuildBaudPart(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex >= 0 Then
        BuildBaudPart = "baud=" & Combo.value
    Else
        BuildBaudPart = ""
    End If
End Function

' -----------------------------------------------------------------------------
' BuildDataBitsPart (Function), source lines 280-286
' -----------------------------------------------------------------------------
Private Function BuildDataBitsPart(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex >= 0 Then
        BuildDataBitsPart = "data=" & Combo.value
    Else
        BuildDataBitsPart = ""
    End If
End Function

' -----------------------------------------------------------------------------
' BuildParityPart (Function), source lines 287-305
' -----------------------------------------------------------------------------
Private Function BuildParityPart(Combo As MSForms.ComboBox) As String
    Dim parityCode As String

    If Combo.ListIndex < 0 Then
        BuildParityPart = ""
        Exit Function
    End If

    Select Case UCase(Combo.value)
        Case "NONE": parityCode = "N"
        Case "ODD": parityCode = "O"
        Case "EVEN": parityCode = "E"
        Case "MARK": parityCode = "M"
        Case "SPACE": parityCode = "S"
        Case Else: parityCode = "N"
    End Select

    BuildParityPart = "parity=" & parityCode
End Function

' -----------------------------------------------------------------------------
' BuildStopBitsPart (Function), source lines 306-312
' -----------------------------------------------------------------------------
Private Function BuildStopBitsPart(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex >= 0 Then
        BuildStopBitsPart = "stop=" & Combo.value
    Else
        BuildStopBitsPart = ""
    End If
End Function

' -----------------------------------------------------------------------------
' BuildFlowControlParts (Function), source lines 313-331
' -----------------------------------------------------------------------------
Private Function BuildFlowControlParts(Combo As MSForms.ComboBox) As String
    If Combo.ListIndex < 0 Then
        BuildFlowControlParts = ""
        Exit Function
    End If

    Select Case UCase(Combo.value)
        Case "NONE"
            BuildFlowControlParts = "xon=off odsr=off dtr=on rts=on"
        Case "XON/XOFF"
            BuildFlowControlParts = "xon=on odsr=off dtr=on rts=on"
        Case "RTS/CTS"
            BuildFlowControlParts = "xon=off odsr=off dtr=on rts=handshake"
        Case "DSR/DTR"
            BuildFlowControlParts = "xon=off odsr=on dtr=handshake rts=on"
        Case Else
            BuildFlowControlParts = "xon=off odsr=off dtr=on rts=on"
    End Select
End Function

' -----------------------------------------------------------------------------
' GetFlowControlListIndex (Function), source lines 332-366
' -----------------------------------------------------------------------------
Private Function GetFlowControlListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim parts As Variant
    Dim xon As String, rts As String, dtr As String, odsr As String
    Dim i As Long

    ' Разбиваем строку настроек по пробелам
    parts = Split(ComSettingsStr, " ")

    ' Ищем нужные параметры
    For i = LBound(parts) To UBound(parts)
        Select Case LCase(left(parts(i), 4))
            Case "xon="
                xon = LCase(Mid(parts(i), 5))
            Case "rts="
                rts = LCase(Mid(parts(i), 5))
            Case "dtr="
                dtr = LCase(Mid(parts(i), 5))
            Case "odsr"
                odsr = LCase(Mid(parts(i), 6))
        End Select
    Next i

    ' По умолчанию — нет управления потоком
    If xon <> "on" And rts <> "handshake" And dtr <> "handshake" And odsr <> "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "None")
    ElseIf xon = "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "Xon/Xoff")
    ElseIf rts = "handshake" Or rts = "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "RTS/CTS")
    ElseIf dtr = "handshake" Or odsr = "on" Then
        GetFlowControlListIndex = ComboBoxIndexOf(Combo, "DSR/DTR")
    Else
        GetFlowControlListIndex = 0 ' не распознано
    End If
End Function

' -----------------------------------------------------------------------------
' ComboBoxIndexOf (Function), source lines 367-376
' -----------------------------------------------------------------------------
Private Function ComboBoxIndexOf(Combo As MSForms.ComboBox, searchText As String) As Long
    Dim i As Long
    For i = 0 To Combo.ListCount - 1
        If StrComp(Combo.List(i), searchText, vbTextCompare) = 0 Then
            ComboBoxIndexOf = i
            Exit Function
        End If
    Next i
    ComboBoxIndexOf = 0
End Function

' -----------------------------------------------------------------------------
' GetStopBitsListIndex (Function), source lines 377-410
' -----------------------------------------------------------------------------
Private Function GetStopBitsListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim stopBitsStr As String
    Dim parts As Variant
    Dim i As Long

    ' Разбиваем строку на отдельные параметры
    parts = Split(ComSettingsStr, " ")

    ' Ищем параметр stop=
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 5)) = "stop=" Then
            stopBitsStr = Mid(parts(i), 6)
            Exit For
        End If
    Next i

    If stopBitsStr = "" Then
        GetStopBitsListIndex = 0 ' stop не найден
        Exit Function
    End If

    ' Нормализуем: "1.5" остаётся строкой, "2" — строка и т.д.
    stopBitsStr = Trim(stopBitsStr)

    ' Ищем в ComboBox
    For i = 0 To Combo.ListCount - 1
        If Combo.List(i) = stopBitsStr Then
            GetStopBitsListIndex = i
            Exit Function
        End If
    Next i

    GetStopBitsListIndex = 0 ' не найдено
End Function

' -----------------------------------------------------------------------------
' GetParityListIndex (Function), source lines 411-451
' -----------------------------------------------------------------------------
Private Function GetParityListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim parityStr As String
    Dim parts As Variant
    Dim i As Long

    ' Разбиваем строку настроек по пробелам
    parts = Split(ComSettingsStr, " ")

    ' Ищем параметр parity=
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 7)) = "parity=" Then
            parityStr = Mid(parts(i), 8)
            Exit For
        End If
    Next i

    If parityStr = "" Then
        GetParityListIndex = 0 ' parity не найден
        Exit Function
    End If

    ' Приводим к удобному формату для сравнения
    Select Case UCase(parityStr)
        Case "N": parityStr = "None"
        Case "O": parityStr = "Odd"
        Case "E": parityStr = "Even"
        Case "M": parityStr = "Mark"
        Case "S": parityStr = "Space"
        ' другие — игнорируем
    End Select

    ' Ищем соответствие в ComboBox
    For i = 0 To Combo.ListCount - 1
        If StrComp(Combo.List(i), parityStr, vbTextCompare) = 0 Then
            GetParityListIndex = i
            Exit Function
        End If
    Next i

    GetParityListIndex = 0 ' не найдено
End Function

' -----------------------------------------------------------------------------
' GetDataBitsListIndex (Function), source lines 452-487
' -----------------------------------------------------------------------------
Private Function GetDataBitsListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim dataBitsStr As String
    Dim dataBits As Long
    Dim parts As Variant
    Dim i As Long

    ' Разбиваем строку настроек по пробелам
    parts = Split(ComSettingsStr, " ")

    ' Ищем параметр data=
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 5)) = "data=" Then
            dataBitsStr = Mid(parts(i), 6)
            Exit For
        End If
    Next i

    If dataBitsStr = "" Then
        GetDataBitsListIndex = 3 ' параметр data= не найден
        Exit Function
    End If

    On Error Resume Next
    dataBits = CLng(dataBitsStr)
    On Error GoTo 0

    ' Сравниваем с элементами ComboBox
    For i = 0 To Combo.ListCount - 1
        If CLng(Combo.List(i)) = dataBits Then
            GetDataBitsListIndex = i
            Exit Function
        End If
    Next i

    GetDataBitsListIndex = 3 ' не найдено
End Function

' -----------------------------------------------------------------------------
' GetBaudListIndex (Function), source lines 488-525
' -----------------------------------------------------------------------------
Private Function GetBaudListIndex(Combo As MSForms.ComboBox, ComSettingsStr As String) As Long
    Dim baudRateStr As String
    Dim baudRate As Long
    Dim parts As Variant
    Dim i As Long

    ' Разбираем строку на части
    parts = Split(ComSettingsStr, " ")
    
    ' Ищем часть, начинающуюся с "baud="
    For i = LBound(parts) To UBound(parts)
        If LCase(left(parts(i), 5)) = "baud=" Then
            baudRateStr = Mid(parts(i), 6) ' пропускаем "baud="
            Exit For
        End If
    Next i

    If baudRateStr = "" Then
        GetBaudListIndex = 5 ' baud не найден
        Exit Function
    End If

    ' Преобразуем в число
    On Error Resume Next
    baudRate = CLng(baudRateStr)
    On Error GoTo 0

    ' Ищем индекс в ComboBox
    For i = 0 To Combo.ListCount - 1
        If CLng(Combo.List(i)) = baudRate Then
            GetBaudListIndex = i
            Exit Function
        End If
    Next i

    ' Если не найдено — вернуть -5
    GetBaudListIndex = 5
End Function

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 526-528
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    isActive = False
End Sub
