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
