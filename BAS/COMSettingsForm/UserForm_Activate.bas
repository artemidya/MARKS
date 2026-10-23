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
