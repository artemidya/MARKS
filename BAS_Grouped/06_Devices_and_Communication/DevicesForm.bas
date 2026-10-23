' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: DevicesForm.frm
' Procedures: 76
' =============================================================================

' -----------------------------------------------------------------------------
' BlocksSocket1_OnStatisticChange (Sub), source lines 33-38
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnStatisticChange(client As BlocksSocket)
    If LoadPlotterProcess = True Then
        Label40.Caption = client.Statistic_PercentReceiving & "%"
        DoEvents
    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton6_Click (Sub), source lines 40-43
' -----------------------------------------------------------------------------
Private Sub CommandButton6_Click()
    TextBox1.text = Format(98.43, "0.00") & "%"
    TextBox2.text = Format(98.43, "0.00") & "%"
End Sub

' -----------------------------------------------------------------------------
' CommandButton7_Click (Sub), source lines 45-48
' -----------------------------------------------------------------------------
Private Sub CommandButton7_Click()
    TextBox1.text = Format(100, "0.00") & "%"
    TextBox2.text = Format(100, "0.00") & "%"
End Sub

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 54-207
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    RefreshLists
    DevicesForm.Caption = localization.ActiveLng.Item("DevicesForm.Caption")
    
    Frame4.ScrollTop = 0
    Combo2.AddItem "HPGL"
    Combo2.AddItem "DMPL"
    Combo2.ListIndex = 0
    
    ComboBox1.AddItem "0;0 " & localization.ActiveLng.Item("DevicesForm.ComboBox1.R")
    ComboBox1.AddItem "0;0 " & localization.ActiveLng.Item("DevicesForm.ComboBox1.L")
    ComboBox1.AddItem "0;0 " & localization.ActiveLng.Item("DevicesForm.ComboBox1.C")
    ComboBox1.ListIndex = 0
    
    ComboBox6.AddItem localization.ActiveLng.Item("DevicesForm.ComboBox6.Abs")
    ComboBox6.AddItem localization.ActiveLng.Item("DevicesForm.ComboBox6.Rel")
    ComboBox6.ListIndex = 0
    
    ComboBox3.AddItem "-90°"
    ComboBox3.AddItem "0°"
    ComboBox3.AddItem "+90°"
    ComboBox3.AddItem "+180°"
    ComboBox3.ListIndex = 1
    
    ComboBox4.AddItem localization.ActiveLng.Item("DevicesForm.ComboBox4.Lines")
'    ComboBox4.AddItem localization.ActiveLng.Item("DevicesForm.ComboBox4.Arcs")
    ComboBox4.ListIndex = 0
    
    
    ComboBox5.AddItem localization.ActiveLng.Item("DevicesForm.ComboBox5.H") '"Âûñîêèé"
    ComboBox5.AddItem localization.ActiveLng.Item("DevicesForm.ComboBox5.M") '"Ñðåäíèé"
    ComboBox5.AddItem localization.ActiveLng.Item("DevicesForm.ComboBox5.L") '"Íèçêèé"
    ComboBox5.ListIndex = 1
    

    With Profiles(PlotterForm.TabStrip1.SelectedItem.index)
        Text4.text = .Device
        Combo2.ListIndex = ListIndexByName(CommandLanguageToString(.CommandLanguage), Combo2)
        ComboBox1.ListIndex = .ZeroPosition
        
        If .CoordinateSystem = Relative Then
            ComboBox6.ListIndex = 1
        Else
            ComboBox6.ListIndex = 0
        End If
        
        CheckBox3.value = .CornerPoints
        TextBox13.text = .MaxSheetWidth & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        TextBox14.text = .MaxSheetHeight & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        ComboBox3.ListIndex = .Turn
        TextBox1.text = Format(Abs(.HeightScale * 100), "0.00") & "%"
        If .HeightScale < 0 Then CheckBox1.value = True
        TextBox2.text = Format(Abs(.WidthScale * 100), "0.00") & "%"
        If .WidthScale < 0 Then CheckBox2.value = True
        'ComboBox4.ListIndex = .LineType
        ComboBox4.ListIndex = 0
        ComboBox5.ListIndex = .ApproximationLevel
        TextBox3.text = .FilePLT
        TextBox4.text = .MemorySize
        TextBox5.text = .Initializer
        TextBox6.text = .BeginWorkFile
        TextBox7.text = .EndWorkFile
        TextBox8.text = .Terminator
        TextBox9.text = .AccelerationsAsString
        TextBox10.text = .SpeedsAsString
        TextBox11.text = .ForcesAsString
        TextBox12.text = .file
        CheckBox5.value = .Laser
        TextBox15.text = .LaserDX & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        TextBox16.text = .LaserDY & " " & localization.ActiveLng.Item("RaskroyForm.mm")
        CheckBox6.value = .SeparateParametrs
        If .CommandLanguage = DMPL Then
            CheckBox7.Enabled = False
            CheckBox7.value = True
        Else
            CheckBox7.value = .OneCommandOnPath
        End If
        TextBox17.text = .PlotterModelString
    End With
    
    
    Frame3.Caption = localization.ActiveLng.Item("DevicesForm.Frame3.Caption") '"Device list"
    Label9.Caption = localization.ActiveLng.Item("DevicesForm.Label9.Caption") '"Highlight your device model in the list below and click the “Use >>” button"
    Command5.Caption = localization.ActiveLng.Item("DevicesForm.Command5.Caption") & " >>"
    Label11.Caption = localization.ActiveLng.Item("DevicesForm.Label11.Caption") '"If your device model is not in this list, then highlight “Standard device with support for HPGL or DMPL commands” and click the “Use >>” button"
    Label12.Caption = localization.ActiveLng.Item("DevicesForm.Label12.Caption") '"The device list is constantly being updated, to add a new model click “Add...” and enter the full model name of the device. The device will be added to the list by a moderator in approximately 0 to 3 days."
    CommandButton1.Caption = localization.ActiveLng.Item("DevicesForm.CommandButton1.Caption") '"Add.."
    Frame5.Caption = localization.ActiveLng.Item("DevicesForm.Frame5.Caption") '"Current device"
    Label13.Caption = localization.ActiveLng.Item("DevicesForm.Label13.Caption") '"Device name"
    Frame6.Caption = localization.ActiveLng.Item("DevicesForm.Frame6.Caption") '"Control"
    Label14.Caption = localization.ActiveLng.Item("DevicesForm.Label14.Caption") '"Command language"
    Label15.Caption = localization.ActiveLng.Item("DevicesForm.Label15.Caption") '"Point of reference"
    Label49.Caption = localization.ActiveLng.Item("DevicesForm.Label49.Caption") '"Coordinate system"
    CheckBox7.Caption = localization.ActiveLng.Item("DevicesForm.CheckBox7.Caption") '"One command per line"
    CheckBox5.Caption = localization.ActiveLng.Item("DevicesForm.CheckBox5.Caption") '"Laser pointer for working with tags"
    Label44.Caption = localization.ActiveLng.Item("DevicesForm.Label44.Caption") '"Distance dX"
    Label45.Caption = localization.ActiveLng.Item("DevicesForm.Label45.Caption") '"Distance dY"
    CheckBox3.Caption = localization.ActiveLng.Item("DevicesForm.CheckBox3.Caption") '"Put points when going around the perimeter at the corners"
    Frame12.Caption = localization.ActiveLng.Item("DevicesForm.Frame12.Caption") '"Working area"
    Label41.Caption = localization.ActiveLng.Item("DevicesForm.Label41.Caption") '"Maximum width"
    Label42.Caption = localization.ActiveLng.Item("DevicesForm.Label42.Caption") '"Maximum height"
    Frame7.Caption = localization.ActiveLng.Item("DevicesForm.Frame7.Caption") '"Transformations"
    Label20.Caption = localization.ActiveLng.Item("DevicesForm.Label20.Caption") '"Turn"
    Label19.Caption = localization.ActiveLng.Item("DevicesForm.Label19.Caption") '"Scale by width"
    Label18.Caption = localization.ActiveLng.Item("DevicesForm.Label18.Caption") '"Scale by height"
    CheckBox2.Caption = localization.ActiveLng.Item("DevicesForm.CheckBox2.Caption") '"Reflect horizontally"
    CheckBox1.Caption = localization.ActiveLng.Item("DevicesForm.CheckBox1.Caption") '"Reflect vertically"
    Frame8.Caption = localization.ActiveLng.Item("DevicesForm.Frame8.Caption") '"Approximation"
    Label21.Caption = localization.ActiveLng.Item("DevicesForm.Label21.Caption") '"Line type"
    Label22.Caption = localization.ActiveLng.Item("DevicesForm.Label22.Caption") '"Level"
    Label23.Caption = localization.ActiveLng.Item("DevicesForm.Label23.Caption") '"If your device does not support arcs, then choose an approximation to a set of straight lines"
    Frame9.Caption = localization.ActiveLng.Item("DevicesForm.Frame9.Caption") '"Output file"
    Label24.Caption = localization.ActiveLng.Item("DevicesForm.Label24.Caption") '"File name"
    Label25.Caption = localization.ActiveLng.Item("DevicesForm.Label25.Caption") '"For output via flash drive"
    Label26.Caption = localization.ActiveLng.Item("DevicesForm.Label26.Caption") '"Memory size, bytes"
    Label27.Caption = localization.ActiveLng.Item("DevicesForm.Label27.Caption") '"If the file size exceeds the specified size, then the file will be split into multiple parts. Set the file size to be no larger than the RAM capacity of the device"
    Label28.Caption = localization.ActiveLng.Item("DevicesForm.Label28.Caption") '"Initialization"
    Label29.Caption = localization.ActiveLng.Item("DevicesForm.Label29.Caption") '"Add to the beginning of a file"
    Label30.Caption = localization.ActiveLng.Item("DevicesForm.Label30.Caption") '"Add to the end of file"
    Label31.Caption = localization.ActiveLng.Item("DevicesForm.Label31.Caption") '"Finalization"
    Label32.Caption = localization.ActiveLng.Item("DevicesForm.Label32.Caption") '"For example, you can add a set of commands to the end of the file that allows you to cut through the film in a single line or rewind the film. Before adding commands, it is better to do a preliminary test and debugging in the console"
    Frame10.Caption = localization.ActiveLng.Item("DevicesForm.Frame10.Caption") '"Hardware parameter lists"
    Label34.Caption = localization.ActiveLng.Item("DevicesForm.Label34.Caption") '"Separate the values of the list items with ';' for example: '100;200;300;400'"
    Label33.Caption = localization.ActiveLng.Item("DevicesForm.Label33.Caption") & ", " & localization.ActiveLng.Item("PlotterForm.m/s") & ChrW(178)
    Label35.Caption = localization.ActiveLng.Item("DevicesForm.Label35.Caption") & ", " & localization.ActiveLng.Item("PlotterForm.mm/s")
    CheckBox6.Caption = localization.ActiveLng.Item("DevicesForm.CheckBox6.Caption") '"Separate value setting for the lowered knife and separate value setting for the lifted knife"
    Label37.Caption = localization.ActiveLng.Item("DevicesForm.Label37.Caption") & ", " & localization.ActiveLng.Item("PlotterForm.g")
    Frame11.Caption = localization.ActiveLng.Item("DevicesForm.Frame11.Caption") '"Settings file for the current device"
    CommandButton3.Caption = localization.ActiveLng.Item("DevicesForm.CommandButton3.Caption") '"Save"
    Label38.Caption = localization.ActiveLng.Item("DevicesForm.Label38.Caption") '"File path"
    CommandButton2.Caption = localization.ActiveLng.Item("DevicesForm.CommandButton2.Caption") '"Show file"
    Label39.Caption = localization.ActiveLng.Item("DevicesForm.Label39.Caption")
    
    Dim q As Long
    Dim vap As String
    With PlotterForm.OutPuts
        Dim t As OutType
        For q = 0 To .count - 1
            t = .OutTypeByIndex(q)
            If t = Interface Then
                vap = UCase(.VidAndPidByIndex(q))
                If vap <> "" Then
                    AddToVidAndPid VidAndPid(), vap
                End If
            End If
        Next
    End With
    
    'Init BlocksSokets
    BlocksSocketShared.HookForm FindWindow(vbNullString, Me.Caption)
    Set BlocksSocket1 = New BlocksSocket
    BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' AddToVidAndPid (Sub), source lines 208-218
' -----------------------------------------------------------------------------
Private Sub AddToVidAndPid(ByRef lVidAndPid() As String, ByVal NewElement As String)
    If Not Not lVidAndPid() Then
        Dim NewIndex As Long
        NewIndex = UBound(lVidAndPid) + 1
        ReDim Preserve lVidAndPid(NewIndex)
        lVidAndPid(NewIndex) = NewElement
    Else
        ReDim lVidAndPid(0)
        lVidAndPid(0) = NewElement
    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton1_Click (Sub), source lines 219-234
' -----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    'AddNewPlotter
    Dim Device As String
    Device = Trim((InputBox(localization.ActiveLng.Item("DevicesForm.InputBox"), "PlotCalc", Profiles(PlotterForm.TabStrip1.SelectedItem.index).ForModeratorDevice)))
    If Device <> "" Then
        Dim buffer1 As New buffer
        With buffer1
            .InterPlatformNetConverter = True
            .size = Len(Device) * 2 + 4
            .bPut_Long Commands.AddPlotter
            .bPut_String Device
            BlocksSocket1.Send .bytes
            Profiles(PlotterForm.TabStrip1.SelectedItem.index).ForModeratorDevice = Device
        End With
    End If
End Sub

' -----------------------------------------------------------------------------
' Label40_DblClick (Sub), source lines 235-239
' -----------------------------------------------------------------------------
Private Sub Label40_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not Connected Then
        BlocksSocket1.Connect UserData.MainAddress, UserData.MainAddressPort
    End If
End Sub

' -----------------------------------------------------------------------------
' RefreshLists (Sub), source lines 240-271
' -----------------------------------------------------------------------------
Private Sub RefreshLists()
    Dim q As Long
    With ListBox1
        .Clear
        .AddItem "    " & localization.ActiveLng.Item("Profile.Device")
        .AddItem "    " & localization.ActiveLng.Item("Profile.DeviceDMPL")
        If Not Not PlottersList Then
            If DeveloperMode Then
                SortPlottersListByDate
                For q = 0 To UBound(PlottersList)
                    .AddItem DateLngToString(PlottersList(q).Date) & " " & PlottersList(q).Device
                Next
            Else
                For q = 0 To UBound(PlottersList)
                    If DeviceIsPresentInMySystem(PlottersList(q)) Then
                        .AddItem ChrW(10003) & PlottersList(q).Device
                    Else
                        .AddItem "    " & PlottersList(q).Device
                    End If
                Next
            End If
        End If
    End With
    With ListBox2
        .Clear
        If Not Not NeedAddList Then
            For q = 0 To UBound(NeedAddList)
                .AddItem DateLngToString(NeedAddList(q).Date) & " " & NeedAddList(q).Device
            Next
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' DeviceIsPresentInMySystem (Function), source lines 272-286
' -----------------------------------------------------------------------------
Private Function DeviceIsPresentInMySystem(pi As PlotterInfo) As Boolean
    If pi.VidAndPid <> "" Then
        Dim q As Long
        If Not Not VidAndPid() Then
            With pi
                For q = 0 To UBound(VidAndPid)
                    If VidAndPid(q) = UCase(.VidAndPid) Then
                        DeviceIsPresentInMySystem = True
                        Exit For
                    End If
                Next
            End With
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' DateLngToString (Function), source lines 287-293
' -----------------------------------------------------------------------------
Private Function DateLngToString(DateLng As Long) As String
    On Error GoTo 10
    Dim t As String
    t = CStr(DateLng)
    DateLngToString = right(t, 2) & "." & Mid(t, 5, 2) & "." & left(t, 4)
10
End Function

' -----------------------------------------------------------------------------
' BlocksSocket1_OnConnectionProcessIsStarted (Sub), source lines 295-297
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnConnectionProcessIsStarted()
    Label40.Caption = localization.ActiveLng.Item("DevicesForm.Label40.Caption_conTo") & " " & UserData.MainAddress & ":" & UserData.MainAddressPort
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnConnect (Sub), source lines 298-303
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnConnect(client As BlocksSocket)
    Label40.Caption = localization.ActiveLng.Item("DevicesForm.Label40.Caption_con")
    Connected = True
    BlocksSocket1.Send GetList
    CommandButton1.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' BlocksSocket1_OnRecive (Sub), source lines 304-366
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' Authorization (Function), source lines 367-387
' -----------------------------------------------------------------------------
Private Function Authorization(Optional Err As Boolean) As Boolean
    If Err = True Then lAuthorization = False
    If Not lAuthorization Then
        Dim pass As String
        pass = InputBox("Enter Moderator password", "Developer Mode", lPass)
        If Len(pass) > 4 Then
            lPass = pass
            'Send
            Dim buffer1 As New buffer
            With buffer1
                .InterPlatformNetConverter = True
                .size = Len(pass) * 2 + 4
                .bPut_Long Commands.Authorization
                .bPut_String pass
                BlocksSocket1.Send .bytes
            End With
            lAuthorization = True
        End If
    End If
    Authorization = lAuthorization
End Function

' -----------------------------------------------------------------------------
' InsertContent (Sub), source lines 388-456
' -----------------------------------------------------------------------------
Private Sub InsertContent(ByRef content As String)
    On Error Resume Next
    Initialized = False
    Text4.text = GetTag("Device", content)
    Combo2.ListIndex = ListIndexByName(GetTag("CommandLanguage", content), Combo2)
    ComboBox1.ListIndex = StringToLong(GetTag("ZeroPosition", content))
    CheckBox3.value = StringToBoolean(GetTag("CornerPoints", content))
    
    Dim sc As String
    sc = GetTag("CoordinateSystem", content)
    If UCase(sc) = "RELATIVE" Then
        ComboBox6.ListIndex = 1
    Else
        ComboBox6.ListIndex = 0
    End If
    
    TextBox13.text = StringToLong(GetTag("MaxSheetWidth", content))
    TextBox14.text = StringToLong(GetTag("MaxSheetHeight", content))
    ComboBox3.ListIndex = StringToLong(GetTag("Turn", content))
    
    Dim HeightScale As Double, WidthScale As Double
    HeightScale = ValPD(GetTag("HeightScale", content))
    WidthScale = ValPD(GetTag("WidthScale", content))
    TextBox1.text = Format(Abs(HeightScale * 100), "0.00") & "%"
    TextBox2.text = Format(Abs(WidthScale * 100), "0.00") & "%"
    If HeightScale < 0 Then CheckBox1.value = True Else CheckBox1.value = False
    If WidthScale < 0 Then CheckBox2.value = True Else CheckBox2.value = False
    
'    ComboBox4.ListIndex = StringToLong(GetTag("LineType", content))
    ComboBox4.ListIndex = 0
    ComboBox5.ListIndex = StringToLong(GetTag("ApproximationLevel", content))
    TextBox3.text = GetTag("FilePLT", content)
    TextBox4.text = StringToLong(GetTag("MemorySize", content))
    TextBox5.text = GetTag("Initializer", content)
    TextBox6.text = GetTag("BeginWorkFile", content)
    TextBox7.text = GetTag("EndWorkFile", content)
    TextBox8.text = GetTag("Terminator", content)
    
    Dim arrStr() As String
    arrStr = Split(GetTag("Accelerations", content), ";")
    TextBox9.text = ""
    If UBound(arrStr) > -1 Then
        Dim q As Long
        For q = 0 To UBound(arrStr) - 1
            TextBox9.text = TextBox9.text & Replace(StringToLong(arrStr(q)) / 10, ",", ".") & ";"
        Next
        TextBox9.text = TextBox9.text & Replace(StringToLong(arrStr(q)) / 10, ",", ".")
    End If
    
    TextBox10.text = GetTag("Speeds", content)
    TextBox11.text = GetTag("Forces", content)
    CheckBox5.value = StringToBoolean(GetTag("Laser", content))
    TextBox15.text = StringToLong(GetTag("LaserDX", content))
    TextBox16.text = StringToLong(GetTag("LaserDY", content))
    CheckBox6.value = StringToBoolean(GetTag("SeparateParametrs", content))
    If Combo2.ListIndex = 1 Then
        CheckBox7.Enabled = False
        CheckBox7.value = True
    Else
        CheckBox7.Enabled = True
        CheckBox7.value = StringToBoolean(GetTag("OneCommandOnPath", content))
    End If
    TextBox13_AfterUpdate
    TextBox14_AfterUpdate
    TextBox15_AfterUpdate
    TextBox16_AfterUpdate
    TextBox17.text = GetTag("PlotterModelString", content)
    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' FindMaxDate (Function), source lines 457-464
' -----------------------------------------------------------------------------
Private Function FindMaxDate(pi() As PlotterInfo) As Long
    If Not Not pi Then
        Dim q As Long
        For q = 0 To UBound(pi)
            If pi(q).Date > FindMaxDate Then FindMaxDate = pi(q).Date
        Next
    End If
End Function

' -----------------------------------------------------------------------------
' BlocksSocket1_OnDisconnect (Sub), source lines 465-473
' -----------------------------------------------------------------------------
Private Sub BlocksSocket1_OnDisconnect(client As BlocksSocket)
    Label40.Caption = localization.ActiveLng.Item("DevicesForm.Label40.Caption_dicnct")
    Connected = False
    Command5.Enabled = False
    CommandButton1.Enabled = False
    Erase PlottersList
    Erase NeedAddList
    RefreshLists
End Sub

' -----------------------------------------------------------------------------
' GetList (Function), source lines 474-480
' -----------------------------------------------------------------------------
Private Function GetList() As Byte()
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.size = 4
    buffer1.bPut_Long Commands.GetPlotterList
    GetList = buffer1.bytes
End Function

' -----------------------------------------------------------------------------
' Text4_Change (Sub), source lines 481-483
' -----------------------------------------------------------------------------
Private Sub Text4_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' Command5_Click (Sub), source lines 484-510
' -----------------------------------------------------------------------------
Private Sub Command5_Click()
    If ListBox1.ListIndex = 0 Then
        If MsgBox(localization.ActiveLng.Item("DevicesForm.MsgBox.Repl"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then
            SetHPGLstandart
            Frame4.ScrollTop = 0
        End If
    ElseIf ListBox1.ListIndex = 1 Then
        If MsgBox(localization.ActiveLng.Item("DevicesForm.MsgBox.Repl"), vbYesNoCancel Or vbQuestion, "PlotCalc") = vbYes Then
            SetDMPLstandart
            Frame4.ScrollTop = 0
        End If
    ElseIf ListBox1.ListIndex > 1 Then
        'Internet
        LoadPlotterProcess = True
        TempLabel40 = Label40.Caption
        Dim ID As Long
        ID = PlottersList(ListBox1.ListIndex - 2).ID
        Dim buffer1 As New buffer
        With buffer1
            .InterPlatformNetConverter = True
            .size = 8
            .bPut_Long Commands.GetPlotter
            .bPut_Long ID
            BlocksSocket1.Send .bytes
        End With
    End If
End Sub

' -----------------------------------------------------------------------------
' SetHPGLstandart (Sub), source lines 511-555
' -----------------------------------------------------------------------------
Private Sub SetHPGLstandart()
    Initialized = False
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device = localization.ActiveLng.Item("Profile.Device")
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage = HPGL
    Text4.text = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device
    Combo2.ListIndex = ListIndexByName("HPGL", Combo2)
    ComboBox1.ListIndex = 0
    ComboBox6.ListIndex = 0
    CheckBox5.value = False
    TextBox15.text = ""
    TextBox16.text = ""
    CheckBox3.value = True
'    CheckBox4.value = False
    TextBox13.text = ""
    TextBox14.text = ""
    ComboBox3.ListIndex = 1
    TextBox1.text = "100"
    TextBox2.text = "100"
    CheckBox1.value = False
    CheckBox2.value = False
    ComboBox4.ListIndex = 0
    ComboBox5.ListIndex = 1
    TextBox3.text = "*.PLT"
    TextBox4.text = "4000000"
    TextBox5.text = PlotterLanguage.Init
    TextBox6.text = ""
    TextBox7.text = ""
    TextBox8.text = PlotterLanguage.TERM
    TextBox9.text = ""
    TextBox10.text = ""
    TextBox11.text = ""
    TextBox17.text = ""
    CheckBox6.value = False
    CheckBox7.value = False
    CheckBox7.Enabled = True
    
    TextBox1_AfterUpdate
    TextBox2_AfterUpdate
    TextBox4_AfterUpdate
    TextBox13_AfterUpdate
    TextBox14_AfterUpdate
    TextBox15_AfterUpdate
    TextBox16_AfterUpdate
    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' SetDMPLstandart (Sub), source lines 556-600
' -----------------------------------------------------------------------------
Private Sub SetDMPLstandart()
    Initialized = False
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device = localization.ActiveLng.Item("Profile.DeviceDMPL")
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage = DMPL
    Text4.text = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device
    Combo2.ListIndex = ListIndexByName("DMPL", Combo2)
    ComboBox1.ListIndex = 0
    ComboBox6.ListIndex = 0
    CheckBox5.value = False
    TextBox15.text = ""
    TextBox16.text = ""
    CheckBox3.value = True
'    CheckBox4.value = False
    TextBox13.text = ""
    TextBox14.text = ""
    ComboBox3.ListIndex = 1
    TextBox1.text = "100"
    TextBox2.text = "100"
    CheckBox1.value = False
    CheckBox2.value = False
    ComboBox4.ListIndex = 0
    ComboBox5.ListIndex = 1
    TextBox3.text = "*.PLT"
    TextBox4.text = "4000000"
    TextBox5.text = PlotterLanguage.Init
    TextBox6.text = ""
    TextBox7.text = ""
    TextBox8.text = PlotterLanguage.TERM
    TextBox9.text = ""
    TextBox10.text = ""
    TextBox11.text = ""
    TextBox17.text = ""
    CheckBox6.value = False
    CheckBox7.value = True
    CheckBox7.Enabled = False
    
    TextBox1_AfterUpdate
    TextBox2_AfterUpdate
    TextBox4_AfterUpdate
    TextBox13_AfterUpdate
    TextBox14_AfterUpdate
    TextBox15_AfterUpdate
    TextBox16_AfterUpdate
    Initialized = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox5_Change (Sub), source lines 601-614
' -----------------------------------------------------------------------------
Private Sub CheckBox5_Change()
    If Initialized Then CommandButton3.Enabled = True
    If CheckBox5.value Then
        Label44.Enabled = True
        Label45.Enabled = True
        TextBox15.Enabled = True
        TextBox16.Enabled = True
    Else
        Label44.Enabled = False
        Label45.Enabled = False
        TextBox15.Enabled = False
        TextBox16.Enabled = False
    End If
End Sub

' -----------------------------------------------------------------------------
' TextBox15_AfterUpdate (Sub), source lines 615-621
' -----------------------------------------------------------------------------
Private Sub TextBox15_AfterUpdate()
    On Error GoTo 10
    TextBox15.text = CLng(val(TextBox15.text)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    Exit Sub
10
    TextBox15.text = Profiles(PlotterForm.TabStrip1.SelectedItem.index).LaserDX & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub

' -----------------------------------------------------------------------------
' TextBox16_AfterUpdate (Sub), source lines 622-628
' -----------------------------------------------------------------------------
Private Sub TextBox16_AfterUpdate()
    On Error GoTo 10
    TextBox16.text = CLng(val(TextBox16.text)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    Exit Sub
10
    TextBox16.text = Profiles(PlotterForm.TabStrip1.SelectedItem.index).LaserDY & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub

' -----------------------------------------------------------------------------
' TextBox15_Change (Sub), source lines 629-631
' -----------------------------------------------------------------------------
Private Sub TextBox15_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox16_Change (Sub), source lines 632-634
' -----------------------------------------------------------------------------
Private Sub TextBox16_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CommandButton2_Click (Sub), source lines 636-640
' -----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    On Error GoTo 10
    ShellExecuteA FindWindow(vbNullString, Me.Caption), "open", "explorer", "/select, " & TextBox12.text, "", 1
10
End Sub

' -----------------------------------------------------------------------------
' ComboBox4_Change (Sub), source lines 641-652
' -----------------------------------------------------------------------------
Private Sub ComboBox4_Change()
    If Initialized Then CommandButton3.Enabled = True
    If ComboBox4.ListIndex = 0 Then
        Label23.Enabled = False
        Label22.Enabled = True
        ComboBox5.Enabled = True
    Else
        Label23.Enabled = True
        Label22.Enabled = False
        ComboBox5.Enabled = False
    End If
End Sub

' -----------------------------------------------------------------------------
' ListBox1_Change (Sub), source lines 653-659
' -----------------------------------------------------------------------------
Private Sub ListBox1_Change()
    If ListBox1.ListIndex < 2 Then
        Command5.Enabled = True
    ElseIf Connected = True Then
        Command5.Enabled = True
    End If
End Sub

' -----------------------------------------------------------------------------
' Combo2_Change (Sub), source lines 660-688
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' CheckBox7_Change (Sub), source lines 689-691
' -----------------------------------------------------------------------------
Private Sub CheckBox7_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox17_Change (Sub), source lines 692-694
' -----------------------------------------------------------------------------
Private Sub TextBox17_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox1_Change (Sub), source lines 695-697
' -----------------------------------------------------------------------------
Private Sub CheckBox1_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox2_Change (Sub), source lines 698-700
' -----------------------------------------------------------------------------
Private Sub CheckBox2_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox3_Change (Sub), source lines 701-703
' -----------------------------------------------------------------------------
Private Sub CheckBox3_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' ComboBox1_Change (Sub), source lines 704-706
' -----------------------------------------------------------------------------
Private Sub ComboBox1_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' ComboBox3_Change (Sub), source lines 707-709
' -----------------------------------------------------------------------------
Private Sub ComboBox3_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' ComboBox5_Change (Sub), source lines 710-712
' -----------------------------------------------------------------------------
Private Sub ComboBox5_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox10_Change (Sub), source lines 714-716
' -----------------------------------------------------------------------------
Private Sub TextBox10_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox11_Change (Sub), source lines 717-719
' -----------------------------------------------------------------------------
Private Sub TextBox11_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox13_Change (Sub), source lines 720-722
' -----------------------------------------------------------------------------
Private Sub TextBox13_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' CheckBox6_Click (Sub), source lines 723-725
' -----------------------------------------------------------------------------
Private Sub CheckBox6_Click()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox14_Change (Sub), source lines 726-728
' -----------------------------------------------------------------------------
Private Sub TextBox14_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox2_Change (Sub), source lines 729-751
' -----------------------------------------------------------------------------
Private Sub TextBox2_Change()
    If Initialized Then CommandButton3.Enabled = True
    If ValP(TextBox1.text) = 100 And ValP(TextBox2.text) = 100 Then
        CommandButton7.Enabled = False
        Label50.Enabled = False
        Image1.Picture = eu_disabled.Picture
    Else
        CommandButton7.Enabled = True
        Label50.Enabled = True
        Image1.Picture = eu_enabled.Picture
    End If
    
    If ValP(TextBox1.text) = 98.43 And ValP(TextBox2.text) = 98.43 Then
        CommandButton6.Enabled = False
        Label51.Enabled = False
        Image2.Picture = us_disabled.Picture
    Else
        CommandButton6.Enabled = True
        Label51.Enabled = True
        Image2.Picture = us_enabled.Picture
    End If

End Sub

' -----------------------------------------------------------------------------
' TextBox1_Change (Sub), source lines 752-773
' -----------------------------------------------------------------------------
Private Sub TextBox1_Change()
    If Initialized Then CommandButton3.Enabled = True
    If ValP(TextBox1.text) = 100 And ValP(TextBox2.text) = 100 Then
        CommandButton7.Enabled = False
        Label50.Enabled = False
        Image1.Picture = eu_disabled.Picture
    Else
        CommandButton7.Enabled = True
        Label50.Enabled = True
        Image1.Picture = eu_enabled.Picture
    End If
    
    If ValP(TextBox1.text) = 98.43 And ValP(TextBox2.text) = 98.43 Then
        CommandButton6.Enabled = False
        Label51.Enabled = False
        Image2.Picture = us_disabled.Picture
    Else
        CommandButton6.Enabled = True
        Label51.Enabled = True
        Image2.Picture = us_enabled.Picture
    End If
End Sub

' -----------------------------------------------------------------------------
' TextBox3_Change (Sub), source lines 774-776
' -----------------------------------------------------------------------------
Private Sub TextBox3_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox4_Change (Sub), source lines 777-779
' -----------------------------------------------------------------------------
Private Sub TextBox4_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox5_Change (Sub), source lines 780-782
' -----------------------------------------------------------------------------
Private Sub TextBox5_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox6_Change (Sub), source lines 783-785
' -----------------------------------------------------------------------------
Private Sub TextBox6_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox7_Change (Sub), source lines 786-788
' -----------------------------------------------------------------------------
Private Sub TextBox7_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox8_Change (Sub), source lines 789-791
' -----------------------------------------------------------------------------
Private Sub TextBox8_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox9_Change (Sub), source lines 792-794
' -----------------------------------------------------------------------------
Private Sub TextBox9_Change()
    If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' ComboBox6_Change (Sub), source lines 795-797
' -----------------------------------------------------------------------------
Private Sub ComboBox6_Change()
 If Initialized Then CommandButton3.Enabled = True
End Sub

' -----------------------------------------------------------------------------
' TextBox9_AfterUpdate (Sub), source lines 803-805
' -----------------------------------------------------------------------------
Private Sub TextBox9_AfterUpdate()
    TextBox9.text = Replace(TextBox9.text, ",", ";")
End Sub

' -----------------------------------------------------------------------------
' TextBox10_AfterUpdate (Sub), source lines 806-808
' -----------------------------------------------------------------------------
Private Sub TextBox10_AfterUpdate()
    TextBox10.text = Replace(TextBox10.text, ",", ";")
End Sub

' -----------------------------------------------------------------------------
' TextBox11_AfterUpdate (Sub), source lines 809-811
' -----------------------------------------------------------------------------
Private Sub TextBox11_AfterUpdate()
    TextBox11.text = Replace(TextBox11.text, ",", ";")
End Sub

' -----------------------------------------------------------------------------
' TextBox1_AfterUpdate (Sub), source lines 812-820
' -----------------------------------------------------------------------------
Private Sub TextBox1_AfterUpdate()
    On Error GoTo 10
    Dim V As Double
    V = ValP(TextBox1.text) / 100
10
    If V > 5 Then V = 5
    If V >= 0 And V < 0.5 Then V = 0.5
    TextBox1.text = Format(V * 100, "0.00") & "%"
End Sub

' -----------------------------------------------------------------------------
' TextBox2_AfterUpdate (Sub), source lines 822-830
' -----------------------------------------------------------------------------
Private Sub TextBox2_AfterUpdate()
    On Error GoTo 10
    Dim V As Double
    V = ValP(TextBox2.text) / 100
10
    If V > 5 Then V = 5
    If V >= 0 And V < 0.5 Then V = 0.5
    TextBox2.text = Format(V * 100, "0.00") & "%"
End Sub

' -----------------------------------------------------------------------------
' TextBox13_AfterUpdate (Sub), source lines 831-838
' -----------------------------------------------------------------------------
Private Sub TextBox13_AfterUpdate()
    On Error GoTo 10
    Dim V As Long
    V = Int(ValP(TextBox13.text))
    If V > 100000 Then V = 100000
10
    TextBox13.text = V & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub

' -----------------------------------------------------------------------------
' TextBox14_AfterUpdate (Sub), source lines 839-846
' -----------------------------------------------------------------------------
Private Sub TextBox14_AfterUpdate()
    On Error GoTo 10
    Dim V As Long
    V = Int(ValP(TextBox14.text))
    If V > 100000 Then V = 100000
10
    TextBox14.text = V & " " & localization.ActiveLng.Item("RaskroyForm.mm")
End Sub

' -----------------------------------------------------------------------------
' TextBox4_AfterUpdate (Sub), source lines 847-853
' -----------------------------------------------------------------------------
Private Sub TextBox4_AfterUpdate()
On Error GoTo 10
    If ValP(TextBox4.text) < 10000 Then TextBox4.text = "10000"
    Exit Sub
10
    TextBox4.text = "10000"
End Sub

' -----------------------------------------------------------------------------
' ValP (Function), source lines 854-856
' -----------------------------------------------------------------------------
Private Function ValP(ByRef txt As String) As Double
    ValP = Abs(val(Replace$(Replace$(txt, "%", ""), ",", ".")))
End Function

' -----------------------------------------------------------------------------
' ValPD (Function), source lines 857-859
' -----------------------------------------------------------------------------
Private Function ValPD(ByRef txt As String) As Double
    ValPD = val(Replace$(Replace$(txt, "%", ""), ",", "."))
End Function

' -----------------------------------------------------------------------------
' DataFromMeToProfile (Sub), source lines 860-907
' -----------------------------------------------------------------------------
Private Sub DataFromMeToProfile()
    With Profiles(PlotterForm.TabStrip1.SelectedItem.index)
        On Error GoTo 10
        .Device = Text4.text
        .CommandLanguage = StringToCommandLanguage(Combo2.text)
        .ZeroPosition = ComboBox1.ListIndex
        .CornerPoints = CheckBox3.value
        .MaxSheetWidth = CLng(ValP(TextBox13.text))
        .MaxSheetHeight = CLng(ValP(TextBox14.text))
        .Turn = ComboBox3.ListIndex

        Dim sYorX As Double
        sYorX = ValP(TextBox1.text) / 100
        If CheckBox1.value Then sYorX = -sYorX
        .HeightScale = sYorX

        sYorX = ValP(TextBox2.text) / 100
        If CheckBox2.value Then sYorX = -sYorX
        .WidthScale = sYorX

        .LineType = ComboBox4.ListIndex
        .ApproximationLevel = ComboBox5.ListIndex
        .FilePLT = TextBox3.text

        .MemorySize = CLng(val(TextBox4.text))
        .Initializer = TextBox5.text
        .BeginWorkFile = TextBox6.text
        .EndWorkFile = TextBox7.text
        .Terminator = TextBox8.text
        
        .AccelerationsAsString = TextBox9.text
        .SpeedsAsString = TextBox10.text
        .ForcesAsString = TextBox11.text
        .Laser = CheckBox5.value
        .LaserDX = CLng(val(TextBox15.text))
        .LaserDY = CLng(val(TextBox16.text))
        .SeparateParametrs = CheckBox6.value
        .OneCommandOnPath = CheckBox7.value
        .PlotterModelString = TextBox17.text

        If ComboBox6.ListIndex = 1 Then
            .CoordinateSystem = Relative
        Else
            .CoordinateSystem = Absolute
        End If
10
    End With
End Sub

' -----------------------------------------------------------------------------
' CommandButton3_Click (Sub), source lines 908-952
' -----------------------------------------------------------------------------
Private Sub CommandButton3_Click()
    CommandButton3.Enabled = False
    DataFromMeToProfile
    Profiles(PlotterForm.TabStrip1.SelectedItem.index).Save
    
    If DeveloperMode Then
        If Authorization Then
            'Send
            Dim content As String
            
            If Len(Trim(Profiles(PlotterForm.TabStrip1.SelectedItem.index).PlotterModelString)) > 10 Then
                'Creating a temporary object for converting equipment image to SVG
                Dim PltDocument1 As New PltDocument
                With PltDocument1
                    .MaxSheetHeight = Profiles(PlotterForm.TabStrip1.SelectedItem.index).MaxSheetHeight * 1000
                    .MaxSheetWidth = Profiles(PlotterForm.TabStrip1.SelectedItem.index).MaxSheetWidth * 1000
                    .PlotterModelString = Profiles(PlotterForm.TabStrip1.SelectedItem.index).PlotterModelString
                    .PlotterName = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Device
                    .ZeroPosition = Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition
                    content = .SVG_Generator
                End With
                
                Dim buffer2 As New buffer
                With buffer2
                    .InterPlatformNetConverter = True
                    .size = Len(content) * 2 + 4
                    .bPut_Long Commands.SetSVG
                    .bPut_String content
                    BlocksSocket1.Send .bytes
                End With
            End If
            
            content = Profiles(PlotterForm.TabStrip1.SelectedItem.index).content
            Dim buffer1 As New buffer
            With buffer1
                .InterPlatformNetConverter = True
                .size = Len(content) * 2 + 4
                .bPut_Long Commands.SetPlotter
                .bPut_String content
                BlocksSocket1.Send .bytes
            End With
            
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' ListBox1_KeyUp (Sub), source lines 953-965
' -----------------------------------------------------------------------------
Private Sub ListBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode.value = 107 And LastKeyCode = 107 Then
        DeveloperMode = True
        TextBox17.Visible = True
        Me.Caption = "Developer Mode"
        Text4.Locked = False
        Text4.BackStyle = fmBackStyleOpaque
        CommandButton3.Caption = "Send"
        Frame13.Visible = True
        RefreshLists
    End If
    LastKeyCode = KeyCode.value
End Sub

' -----------------------------------------------------------------------------
' GetPlottersList (Function), source lines 966-993
' -----------------------------------------------------------------------------
Private Function GetPlottersList(ByRef content As String) As PlotterInfo()
      Dim content2 As String
      content2 = GetTag("PlottersList", content)
      Dim arr() As String
      Dim arr2() As String
      arr = Split(content2, "<br>", , vbTextCompare)
      Dim lGetPlottersList() As PlotterInfo
      
      If UBound(arr) > -1 Then
            ReDim lGetPlottersList(UBound(arr))
      End If
      
      Dim q As Long
      For q = 0 To UBound(arr)
            arr2() = Split(arr(q), ";", , vbTextCompare)
            If UBound(arr2) > 1 Then
                  With lGetPlottersList(q)
                        On Error GoTo 5
                        .Device = arr2(0)
                        .ID = CLng(val(arr2(1)))
                        .Date = CLng(val(arr2(2)))
                        .VidAndPid = arr2(13)
5
                  End With
            End If
      Next
      GetPlottersList = lGetPlottersList
End Function

' -----------------------------------------------------------------------------
' GetTag (Function), source lines 994-1003
' -----------------------------------------------------------------------------
Private Function GetTag(TagName As String, ByRef FromText As String) As String
      Dim s As String, e As String
      s = InStr(1, FromText, "<" & TagName & ">", vbTextCompare)
      If s <> 0 Then
            e = InStr(s + Len(TagName) + 2, FromText, "</" & TagName & ">", vbTextCompare)
            If e <> 0 Then
                  GetTag = Mid(FromText, s + Len(TagName) + 2, e - s - Len(TagName) - 2)
            End If
      End If
End Function

' -----------------------------------------------------------------------------
' GetNeedAddList (Function), source lines 1004-1030
' -----------------------------------------------------------------------------
Private Function GetNeedAddList(ByRef content As String) As PlotterInfo()
      Dim content2 As String
      content2 = GetTag("NeedAddList", content)
      Dim arr() As String
      Dim arr2() As String
      arr = Split(content2, "<br>", , vbTextCompare)
      Dim lGetNeedAddList() As PlotterInfo
            
      If UBound(arr) > -1 Then
            ReDim lGetNeedAddList(UBound(arr))
      End If
      
      Dim q As Long
      For q = 0 To UBound(arr)
            arr2() = Split(arr(q), ";", , vbTextCompare)
            If UBound(arr2) > 1 Then
                  With lGetNeedAddList(q)
                        On Error GoTo 5
                        .Device = arr2(0)
                        .ID = CLng(val(arr2(1)))
                        .Date = CLng(val(arr2(2)))
5
                  End With
            End If
      Next
      GetNeedAddList = lGetNeedAddList
End Function

' -----------------------------------------------------------------------------
' SortPlottersListByDate (Sub), source lines 1031-1048
' -----------------------------------------------------------------------------
Private Sub SortPlottersListByDate()
    If Not Not PlottersList Then
        Dim i As Long
        Dim fSwitched As Boolean
        Dim plotterTemp As PlotterInfo
        Do
            fSwitched = False
            For i = 0 To UBound(PlottersList) - 1
                If PlottersList(i).Date < PlottersList(i + 1).Date Then
                    fSwitched = True
                    plotterTemp = PlottersList(i)
                    PlottersList(i) = PlottersList(i + 1)
                    PlottersList(i + 1) = plotterTemp
                End If
            Next
        Loop While fSwitched
    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton4_Click (Sub), source lines 1049-1081
' -----------------------------------------------------------------------------
Private Sub CommandButton4_Click()
    
    Dim buffer1 As New buffer
    If ListFocusNumber = 2 Then
        If ListBox2.ListIndex > -1 Then
            If Authorization Then
                'Reject from NeedAddList
                With buffer1
                    .InterPlatformNetConverter = True
                    .size = Len(NeedAddList(ListBox2.ListIndex).Device) * 2 + 4
                    .bPut_Long Commands.Reject
                    .bPut_String NeedAddList(ListBox2.ListIndex).Device
                    BlocksSocket1.Send .bytes
                End With
            End If
        End If
    Else
        If ListBox1.ListIndex > 1 Then
            If MsgBox("Delete this plotter?", vbYesNoCancel Or vbQuestion) = vbYes Then
                If Authorization Then
                    'Delete Plotter
                    With buffer1
                        .InterPlatformNetConverter = True
                        .size = 8
                        .bPut_Long Commands.DeletePlotter
                        .bPut_Long PlottersList(ListBox1.ListIndex - 2).ID
                        BlocksSocket1.Send .bytes
                    End With
                End If
            End If
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' CommandButton5_Click (Sub), source lines 1083-1094
' -----------------------------------------------------------------------------
Private Sub CommandButton5_Click()
    Dim clipboard1 As New ClipBoard
    If ListFocusNumber = 2 Then
        If ListBox2.ListIndex > -1 Then
            clipboard1.SetData NeedAddList(ListBox2.ListIndex).Device
        End If
    Else
        If ListBox1.ListIndex > 1 Then
            clipboard1.SetData PlottersList(ListBox1.ListIndex - 2).Device
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' ListBox1_Click (Sub), source lines 1095-1097
' -----------------------------------------------------------------------------
Private Sub ListBox1_Click()
    ListFocusNumber = 1
End Sub

' -----------------------------------------------------------------------------
' ListBox2_Click (Sub), source lines 1098-1100
' -----------------------------------------------------------------------------
Private Sub ListBox2_Click()
    ListFocusNumber = 2
End Sub

' -----------------------------------------------------------------------------
' UserForm_Terminate (Sub), source lines 1101-1105
' -----------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    BlocksSocket1.Disconnect
    DataFromMeToProfile
    PlotterForm.RefreshFields
End Sub
