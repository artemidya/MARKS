' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: USB_Events.cls
' Procedures: 6
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 77-105
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    hWndForm = GetForegroundWindow
    HookWindow

    'Interfaces
    Dim filter As DEV_BROADCAST_DEVICEINTERFACE
    Dim guid As guid
    guid.Data1 = &HA5DCBF10
    guid.Data2 = &H6530
    guid.Data3 = &H11D2
    guid.Data4(0) = &H90: guid.Data4(1) = &H1F: guid.Data4(2) = &H0
    guid.Data4(3) = &HC0: guid.Data4(4) = &H4F: guid.Data4(5) = &HB9
    guid.Data4(6) = &H51: guid.Data4(7) = &HED

    filter.dbcc_size = LenB(filter)
    filter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
    filter.dbcc_classguid = guid

    'Ports
    Dim filter2 As DEV_BROADCAST_PORT
    filter2.dbcc_size = LenB(filter2)
    filter2.dbcc_devicetype = DBT_DEVTYP_PORT
    filter2.dbcc_reserved = 0
    
    notifyHandle = RegisterDeviceNotification(hWndForm, filter, DEVICE_NOTIFY_WINDOW_HANDLE)
    notifyHandle2 = RegisterDeviceNotification(hWndForm, filter2, DEVICE_NOTIFY_WINDOW_HANDLE)
    
    Set USB_EventsShared.Parent = Me
End Sub

' -----------------------------------------------------------------------------
' Terminate (Sub), source lines 107-109
' -----------------------------------------------------------------------------
Public Sub Terminate()
    Set USB_EventsShared.Parent = Nothing
End Sub

' -----------------------------------------------------------------------------
' HookWindow (Sub), source lines 111-114
' -----------------------------------------------------------------------------
Private Sub HookWindow()
    oldWndProc = GetWindowLongPtr(hWndForm, GWL_WNDPROC)
    SetWindowLongPtr hWndForm, GWL_WNDPROC, AddressOf USB_EventsShared.WindowProc
End Sub

' -----------------------------------------------------------------------------
' UnhookWindow (Sub), source lines 116-120
' -----------------------------------------------------------------------------
Private Sub UnhookWindow()
    If oldWndProc <> 0 Then
        SetWindowLongPtr hWndForm, GWL_WNDPROC, oldWndProc
    End If
End Sub

' -----------------------------------------------------------------------------
' USB_change_check_loop (Sub), source lines 121-126
' -----------------------------------------------------------------------------
Public Sub USB_change_check_loop()
    If USB_EventsShared.Changed And (Timer - USB_EventsShared.lastWorkingTime) > 0.6 Then
        USB_EventsShared.Changed = False
        RaiseEvent Change
    End If
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 128-136
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    If notifyHandle <> 0 Then
        UnregisterDeviceNotification notifyHandle
    End If
    If notifyHandle2 <> 0 Then
        UnregisterDeviceNotification notifyHandle2
    End If
    UnhookWindow
End Sub
