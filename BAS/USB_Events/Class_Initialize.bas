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
