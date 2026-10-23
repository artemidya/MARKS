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
