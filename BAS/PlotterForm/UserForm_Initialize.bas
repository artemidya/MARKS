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
