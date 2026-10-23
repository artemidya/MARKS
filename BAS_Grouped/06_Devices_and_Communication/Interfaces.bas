' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: Interfaces.cls
' Procedures: 33
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 179-181
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    Refresh
End Sub

' -----------------------------------------------------------------------------
' isComOrLPT_Adapter (PropertyGet), source lines 182-184
' -----------------------------------------------------------------------------
Property Get isComOrLPT_Adapter(ByVal i As Long) As Boolean
    isComOrLPT_Adapter = Interfaces(i).isComPortAdapter
End Property

' -----------------------------------------------------------------------------
' PortName (PropertyGet), source lines 185-187
' -----------------------------------------------------------------------------
Property Get PortName(ByVal i As Long) As String
    PortName = Interfaces(i).PortName
End Property

' -----------------------------------------------------------------------------
' Connected (PropertyGet), source lines 188-190
' -----------------------------------------------------------------------------
Property Get Connected(ByVal i As Long) As Boolean
    Connected = Interfaces(i).Active
End Property

' -----------------------------------------------------------------------------
' ExparedID (PropertyGet), source lines 191-200
' -----------------------------------------------------------------------------
Property Get ExparedID(OutPutName As String) As String
    Dim index As Long
    index = GetIndex(OutPutName)

    With Interfaces(index)
        If Not .isComPortAdapter Then
            ExparedID = .ExparedID
        End If
    End With
End Property

' -----------------------------------------------------------------------------
' ExparedIDByIndex (PropertyGet), source lines 201-207
' -----------------------------------------------------------------------------
Property Get ExparedIDByIndex(ByVal index As Long) As String
    With Interfaces(index)
        If Not .isComPortAdapter Then
            ExparedIDByIndex = .ExparedID
        End If
    End With
End Property

' -----------------------------------------------------------------------------
' VidAndPid (PropertyGet), source lines 209-220
' -----------------------------------------------------------------------------
Property Get VidAndPid(OutPutName As String) As String
    Dim index As Long
    index = GetIndex(OutPutName)

    With Interfaces(index)
        If Not .isComPortAdapter Then
            VidAndPid = .VidAndPid
        Else 'if com-port - need check vidAndpid from usb interfaces
            VidAndPid = CheckCOMVidAndPid(.VidAndPid)
        End If
    End With
End Property

' -----------------------------------------------------------------------------
' VidAndPidByIndex (PropertyGet), source lines 221-229
' -----------------------------------------------------------------------------
Property Get VidAndPidByIndex(ByVal index As Long) As String
    With Interfaces(index)
        If Not .isComPortAdapter Then
            VidAndPidByIndex = .VidAndPid
        Else 'if com-port - need check vidAndpid from usb interfaces
            VidAndPidByIndex = CheckCOMVidAndPid(.VidAndPid)
        End If
    End With
End Property

' -----------------------------------------------------------------------------
' CheckCOMVidAndPid (Function), source lines 230-242
' -----------------------------------------------------------------------------
Private Function CheckCOMVidAndPid(VidAndPid As String) As String
    Dim q As Long
    For q = 0 To UBound(Interfaces)
        With Interfaces(q)
            If Not .isComPortAdapter Then
                If UCase(VidAndPid) = UCase(.VidAndPid) Then
                    CheckCOMVidAndPid = VidAndPid
                    Exit Function
                End If
            End If
        End With
    Next
End Function

' -----------------------------------------------------------------------------
' count (PropertyGet), source lines 243-245
' -----------------------------------------------------------------------------
Public Property Get count() As Long
    count = lCount
End Property

' -----------------------------------------------------------------------------
' Names (PropertyGet), source lines 247-249
' -----------------------------------------------------------------------------
Public Property Get Names(i As Long) As String
    Names = Interfaces(i).name
End Property

' -----------------------------------------------------------------------------
' Descriptions (PropertyGet), source lines 251-253
' -----------------------------------------------------------------------------
Public Property Get Descriptions(i As Long) As String
    Descriptions = Interfaces(i).Description
End Property

' -----------------------------------------------------------------------------
' x (Function), source lines 255-259
' -----------------------------------------------------------------------------
Public Function x(OutPutName As String) As Long
    Dim i As Long
    i = GetIndex(OutPutName)
    x = Interfaces(i).x
End Function

' -----------------------------------------------------------------------------
' y (Function), source lines 261-265
' -----------------------------------------------------------------------------
Public Function y(OutPutName As String) As Long
    Dim i As Long
    i = GetIndex(OutPutName)
    y = Interfaces(i).y
End Function

' -----------------------------------------------------------------------------
' ComSettingsStr (PropertyGet), source lines 266-276
' -----------------------------------------------------------------------------
Property Get ComSettingsStr(OutPutName As String) As String
    On Error GoTo 10
    Dim i As Long:    i = GetIndex(OutPutName)
    Dim RegData1 As New RegData
    With Interfaces(i)
        If .isComPortAdapter And InStr(1, .PortName, "COM", vbTextCompare) Then
            ComSettingsStr = RegData1.ComSettingsStr(RemoveTrailingColon(Trim(.PortName)))
        End If
    End With
10
End Property

' -----------------------------------------------------------------------------
' Out (Sub), source lines 277-314
' -----------------------------------------------------------------------------
Public Sub Out(OutPutName As String, ByRef bytes() As Byte, x As Long, y As Long, Optional dntInitialize As Boolean = False)
    Dim i As Long
    i = GetIndex(OutPutName)
    With Interfaces(i)
    
        If (.Init <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer _
        Or .TERM <> Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator) And Not dntInitialize Then
            If .Initialized Then DeInitialize (i)
            Initialize (i)
        End If
        .x = x
        .y = y
        
        If Not Not bytes() Then
            Dim usbHandle As Long
            Dim dataToSend As String
            Dim BytesWritten As Long
        
            If .isComPortAdapter And InStr(1, Interfaces(i).PortName, "COM", vbTextCompare) Then
                'Reset COM
                COM.CommOpen 1, .FullDevicePath, RemoveTrailingColon(Trim(.PortName))
                COM.CommClose 1
            End If
            ' Открываем USB-порт
            usbHandle = CreateFile(.FullDevicePath, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL + FILE_FLAG_SEQUENTIAL_SCAN, 0)
            
            If usbHandle <> -1 Then
                ' Отправляем данные
                Dim size As Long
                size = UBound(bytes()) + 1
                WriteFile usbHandle, bytes(0), size, BytesWritten, 0
                CloseHandle usbHandle
            Else
        '        MsgBox "Ошибка открытия USB-порта!", vbCritical
            End If
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' Initialize (Sub), source lines 315-350
' -----------------------------------------------------------------------------
Private Sub Initialize(i As Long)

    With Interfaces(i)
        .Init = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Initializer
        .TERM = Profiles(PlotterForm.TabStrip1.SelectedItem.index).Terminator
        .Initialized = True
        
        PlotterForm.AddLog .Init
        Dim bytes() As Byte
        bytes = StrToAscii(.Init)
        
        If Not Not bytes() Then
            Dim usbHandle As Long
            Dim dataToSend As String
            Dim BytesWritten As Long
        
            If .isComPortAdapter And InStr(1, Interfaces(i).PortName, "COM", vbTextCompare) Then
                'Reset COM
                COM.CommOpen 1, .FullDevicePath, RemoveTrailingColon(Trim(.PortName))
                COM.CommClose 1
            End If
            ' Открываем USB-порт
            usbHandle = CreateFile(.FullDevicePath, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL + FILE_FLAG_SEQUENTIAL_SCAN, 0)
            
            If usbHandle <> -1 Then
                ' Отправляем данные
                Dim size As Long
                size = UBound(bytes()) + 1
                WriteFile usbHandle, bytes(0), size, BytesWritten, 0
                CloseHandle usbHandle
            Else
        '        MsgBox "Ошибка открытия USB-порта!", vbCritical
            End If
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' DeInitialize (Sub), source lines 351-385
' -----------------------------------------------------------------------------
Private Sub DeInitialize(i As Long)
    With Interfaces(i)
        PlotterForm.AddLog .TERM
        Dim bytes() As Byte
        bytes = StrToAscii(.TERM)
        
        .Initialized = False
        .Init = ""
        .TERM = ""

        If Not Not bytes() Then
            Dim usbHandle As Long
            Dim dataToSend As String
            Dim BytesWritten As Long
        
            If .isComPortAdapter And InStr(1, Interfaces(i).PortName, "COM", vbTextCompare) Then
                'Reset COM
                COM.CommOpen 1, .FullDevicePath, RemoveTrailingColon(Trim(.PortName))
                COM.CommClose 1
            End If
            ' Открываем USB-порт
            usbHandle = CreateFile(.FullDevicePath, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL + FILE_FLAG_SEQUENTIAL_SCAN, 0)
            
            If usbHandle <> -1 Then
                ' Отправляем данные
                Dim size As Long
                size = UBound(bytes()) + 1
                WriteFile usbHandle, bytes(0), size, BytesWritten, 0
                CloseHandle usbHandle
            Else
        '        MsgBox "Ошибка открытия USB-порта!", vbCritical
            End If
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' GetIndex (Function), source lines 386-394
' -----------------------------------------------------------------------------
Private Function GetIndex(OutPutName As String) As Long
    Dim q As Long
    For q = 0 To UBound(Interfaces)
        If Interfaces(q).name = OutPutName Then
            GetIndex = q
            Exit Function
        End If
    Next
End Function

' -----------------------------------------------------------------------------
' Refresh (Sub), source lines 395-476
' -----------------------------------------------------------------------------
Public Sub Refresh()
    Class_Terminate 'Deinitialize
    Erase Interfaces()
    lCount = 0
 
    'USB - {a5dcbf10-6530-11d2-901f-00c04fb951ed}
    Dim USB_GUID As guid
    With USB_GUID
        .Data1 = &HA5DCBF10
        .Data2 = &H6530
        .Data3 = &H11D2
        .Data4(0) = &H90
        .Data4(1) = &H1F
        .Data4(2) = &H0
        .Data4(3) = &HC0
        .Data4(4) = &H4F
        .Data4(5) = &HB9
        .Data4(6) = &H51
        .Data4(7) = &HED
    End With
    
    'WireFire {28d78fad-5a12-11d1-ae5b-0000f803a8c2}
    Dim WireFire_GUID As guid
    With WireFire_GUID
        .Data1 = &H28D78FAD
        .Data2 = &H5A12
        .Data3 = &H11D1
        .Data4(0) = &HAE
        .Data4(1) = &H5B
        .Data4(2) = &H0
        .Data4(3) = &H0
        .Data4(4) = &HF8
        .Data4(5) = &H3
        .Data4(6) = &HA8
        .Data4(7) = &HC2
    End With
    
    'Bluetooth
    '{e0cbf06c-cd8b-4647-bb8a-263b43f0f974}
    Dim Bluetooth_GUID As guid
    With Bluetooth_GUID
        .Data1 = &HE0CBF06C
        .Data2 = &HCD8B
        .Data3 = &H4647
        .Data4(0) = &HBB
        .Data4(1) = &H8A
        .Data4(2) = &H26
        .Data4(3) = &H3B
        .Data4(4) = &H43
        .Data4(5) = &HF0
        .Data4(6) = &HF9
        .Data4(7) = &H74
    End With
    
    'COM 86E0D1E0-8089-11D0-9CE4-08003E301F73
    Dim Com_GUID As guid
    With Com_GUID
        .Data1 = &H86E0D1E0
        .Data2 = &H8089
        .Data3 = &H11D0
        .Data4(0) = &H9C
        .Data4(1) = &HE4
        .Data4(2) = &H8
        .Data4(3) = &H0
        .Data4(4) = &H3E
        .Data4(5) = &H30
        .Data4(6) = &H1F
        .Data4(7) = &H73
    End With
  
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("USB: ", USB_GUID, "HIDClass,WPD,Bluetooth,Net,Biometric")
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("Bluetooth: ", Bluetooth_GUID)
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("WireFire: ", WireFire_GUID)
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("COM: ", Com_GUID)
    

    If Not Not Interfaces() Then
        lCount = UBound(Interfaces) + 1
    Else
        lCount = 0
    End If
End Sub

' -----------------------------------------------------------------------------
' GetDeviceInterfacesListForPlotter (Function), source lines 477-597
' -----------------------------------------------------------------------------
Private Function GetDeviceInterfacesListForPlotter(ClassNameStr As String, ByRef PrnGUID As guid, Optional ByVal filter As String = "") As deviceInfo()
    filter = UCase(filter)
    Dim PnpHandle As LongPtr
    Dim DeviceInterfaceData  As Device_Interface_Data
    Dim FunctionClassDeviceData As Device_Interface_Detail
    Dim ClassSTR As String
    
    Dim d As LongPtr
    If LenB(d) = 8 Then
        'In VBA 64-bit cbSize = 8
        FunctionClassDeviceData.cbSize = 8
    Else
        'In VBA 32-bit cbSize = 5
        FunctionClassDeviceData.cbSize = 5
    End If

    Dim BytesReturned As Long
    Dim deviceName As String, FriendlyName As String, Description As String, MFG As String
    Dim regType As Long
    Dim reqSize As Long
    Dim buffer(0 To 255) As Byte
    Dim devInfo As SP_DEVINFO_DATA
    devInfo.cbSize = LenB(devInfo)
    
    Dim isActive As Boolean, isDefault As Boolean, isRemoved As Boolean
    Dim curInterface As deviceInfo
    Dim arrInterfaces() As deviceInfo
    
    'Get interfaces
    PnpHandle = SetupDiGetClassDevs(PrnGUID, 0, 0, DIGCF_DEVICEINTERFACE)
'     PnpHandle = SetupDiGetClassDevs(PrnGUID, 0, 0, DIGCF_ALLCLASSES)
   '
   
    If (PnpHandle <> -1) Then
        'Get Device Names
        DeviceInterfaceData.cbSize = LenB(DeviceInterfaceData)
        Dim DevIndex As Long
        
        Do While SetupDiEnumDeviceInterfaces(PnpHandle, 0, ByVal VarPtr(PrnGUID), DevIndex, ByVal VarPtr(DeviceInterfaceData))
            If SetupDiGetDeviceInterfaceDetail(PnpHandle, DeviceInterfaceData, FunctionClassDeviceData, UBound(FunctionClassDeviceData.DataPath), BytesReturned, devInfo) Then
            
                'Get Device Name
                deviceName = left$(StrConv(FunctionClassDeviceData.DataPath(), vbUnicode), BytesReturned - 5)
                
                'Get Description
                If SetupDiGetDeviceRegistryProperty(PnpHandle, devInfo, SPDRP_DEVICEDESC, regType, VarPtr(buffer(0)), 256, reqSize) Then
                    Description = left$(StrConv(buffer(), vbUnicode), reqSize - 1)
                Else
                    Description = "" 'no Description
                End If
                
                'Get MFG
                If SetupDiGetDeviceRegistryProperty(PnpHandle, devInfo, SPDRP_MFG, regType, VarPtr(buffer(0)), 256, reqSize) Then
                    MFG = left$(StrConv(buffer(), vbUnicode), reqSize - 1)
                Else
                    MFG = "" 'no MFG
                End If
                
                'Get Friendly Name
                If SetupDiGetDeviceRegistryProperty(PnpHandle, devInfo, SPDRP_FRIENDLYNAME, regType, VarPtr(buffer(0)), 256, reqSize) Then
                    FriendlyName = left$(StrConv(buffer(), vbUnicode), reqSize - 1)
                Else
                    FriendlyName = "" 'no name
                End If
                
                
                'Get Device class SPDRP_CLASS , SPDRP_CLASSGUID , SPDRP_ENUMERATOR_NAME
                If SetupDiGetDeviceRegistryProperty(PnpHandle, devInfo, SPDRP_CLASS, regType, VarPtr(buffer(0)), 256, reqSize) Then
                    ClassSTR = left$(StrConv(buffer(), vbUnicode), reqSize - 1)
                Else
                    ClassSTR = ""
                End If
                
                
                'Active or Not
                If HasFlag(DeviceInterfaceData.Flags, SPINT_ACTIVE) Then isActive = True Else isActive = False
                If HasFlag(DeviceInterfaceData.Flags, SPINT_DEFAULT) Then isDefault = True Else isDefault = False
                If HasFlag(DeviceInterfaceData.Flags, SPINT_REMOVED) Then isRemoved = True Else isRemoved = False
                
                'Erase curInterface
                With curInterface
                    If Trim(FriendlyName) <> "" Then
                        .name = ClassNameStr & FriendlyName
                    Else
                        .name = ClassNameStr & ShotDevicePath(deviceName)
                    End If

                    If InStr(1, ClassSTR, "PORTS", vbTextCompare) > 0 Then
                        'Also need check is this port LPT port (isLptPortAdapterFunc), else: .isComPortAdapter = True
                        .isComPortAdapter = True
                        .PortName = GetPortNameFromDevice(PnpHandle, devInfo)
                        
                    Else
                        .isComPortAdapter = False
                        .PortName = ""
                    End If
                    
                    .Active = isActive
                    .FullDevicePath = deviceName
                    .ExparedID = ExtractExparedId(deviceName) 'for Chinace plotters
                    .VidAndPid = ExtractVidPid(deviceName)
                    .FriendlyName = FriendlyName
                    .Description = Description
                    .MFG = MFG
                End With
                'Filtr
                If InStr(1, filter, ClassSTR, vbTextCompare) <= 0 Then
                    AddToArray arrInterfaces, curInterface
                End If
                
                'Debug.Print FriendlyName & " (" & Description & ") - " & DeviceName & " SPINT_ACTIVE=" & isActive & " SPINT_DEFAULT=" & isDefault & " SPINT_REMOVED=" & isRemoved
                'Debug.Print ""
            End If
            
            DevIndex = DevIndex + 1
        Loop
    End If
    SetupDiDestroyDeviceInfoList (PnpHandle)
    Sort arrInterfaces()
    GetDeviceInterfacesListForPlotter = arrInterfaces
End Function

' -----------------------------------------------------------------------------
' GetPortNameFromDevice (Function), source lines 598-616
' -----------------------------------------------------------------------------
Private Function GetPortNameFromDevice(hDevInfo As LongPtr, devInfoData As SP_DEVINFO_DATA) As String
    Dim hKey As LongPtr
    Dim sBuffer As String * 256
    Dim dwSize As Long
    Dim dwType As Long
    Dim result As Long


    hKey = SetupDiOpenDevRegKey(hDevInfo, devInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ)
    If hKey = 0 Then Exit Function

    dwSize = Len(sBuffer)
    result = RegQueryValueExA(hKey, "PortName", 0, dwType, sBuffer, dwSize)
    If result = 0 And dwType = REG_SZ Then
        GetPortNameFromDevice = left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If

    RegCloseKey hKey
End Function

' -----------------------------------------------------------------------------
' Sort (Sub), source lines 635-652
' -----------------------------------------------------------------------------
Private Sub Sort(arr() As deviceInfo)
    If Not Not arr Then
        Dim i As Long
        Dim fSwitched As Boolean
        Dim arrTemp As deviceInfo
        Do
            fSwitched = False
            For i = 0 To UBound(arr) - 1
                If addZerroForDigits(arr(i).name) > addZerroForDigits(arr(i + 1).name) Then
                    fSwitched = True
                    arrTemp = arr(i)
                    arr(i) = arr(i + 1)
                    arr(i + 1) = arrTemp
                End If
            Next
        Loop While fSwitched
    End If
End Sub

' -----------------------------------------------------------------------------
' addZerroForDigits (Function), source lines 653-713
' -----------------------------------------------------------------------------
Private Function addZerroForDigits(ByVal txt As String, Optional ByVal maxDigits As Long = 50) As String
    txt = Replace(txt, ",", ".")
    Dim q As Long
    Dim code As Integer
    Dim lastCode As Integer
    Dim isStartingDigit As Boolean
    Dim startPositionDigits() As Long
    Dim endPositionDigits() As Long
    Dim digitsCount As Long
    
    Dim toRight() As Boolean ' (???? ?????, ?? ???????? ???? ??????)
    
    For q = 1 To Len(txt)
        code = Asc(Mid(txt, q, 1))
        If (code >= 48 And code <= 57) Then '..  Or code = 46
            If isStartingDigit = False Then
                isStartingDigit = True
                digitsCount = digitsCount + 1
                ReDim Preserve startPositionDigits(digitsCount - 1)
                ReDim Preserve endPositionDigits(digitsCount - 1)
                ReDim Preserve toRight(digitsCount - 1)
                startPositionDigits(digitsCount - 1) = q
                If lastCode = 46 Then toRight(digitsCount - 1) = True
            End If
        Else
            If isStartingDigit = True Then
                isStartingDigit = False
                endPositionDigits(digitsCount - 1) = q
            End If
        End If
        lastCode = code
    Next
    
    If isStartingDigit = True Then endPositionDigits(digitsCount - 1) = q
    
    Dim s As Long: s = 1
    Dim l As Long
    Dim l2 As Long
    Dim newDigits As String
    Dim q2 As Long

    For q = 0 To digitsCount - 1
        l = startPositionDigits(q) - s
        addZerroForDigits = addZerroForDigits & Mid(txt, s, l)
        l2 = endPositionDigits(q) - startPositionDigits(q)
        s = s + l + l2
        
        newDigits = ""
        For q2 = 1 To maxDigits - l2
            newDigits = newDigits & "0"
        Next
        If toRight(q) = True Then
            newDigits = Mid(txt, startPositionDigits(q), l2) & newDigits
        Else
            newDigits = newDigits & Mid(txt, startPositionDigits(q), l2)
        End If
        addZerroForDigits = addZerroForDigits & newDigits
    Next
    l = Len(txt) - s + 1
    addZerroForDigits = addZerroForDigits & Mid(txt, s, l)
End Function

' -----------------------------------------------------------------------------
' AddArrayToArray (Sub), source lines 714-733
' -----------------------------------------------------------------------------
Private Sub AddArrayToArray(ByRef arr() As deviceInfo, arr2() As deviceInfo)

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
' AddToArray (Sub), source lines 734-744
' -----------------------------------------------------------------------------
Private Sub AddToArray(ByRef arr() As deviceInfo, newItem As deviceInfo)
    Dim NewIndex As Long
    If Not Not arr() Then
        NewIndex = UBound(arr) + 1
        ReDim Preserve arr(NewIndex)
        arr(NewIndex) = newItem
    Else
        ReDim arr(0)
        arr(0) = newItem
    End If
End Sub

' -----------------------------------------------------------------------------
' ShotDevicePath (Function), source lines 746-756
' -----------------------------------------------------------------------------
Private Function ShotDevicePath(ByVal path As String) As String
    On Error GoTo 10
    path = UCase(path)
    path = Replace(path, "\\?\", "", , 1)
    path = Replace(path, "\\.\", "", , 1)
    path = Replace(path, "#", "\")

    path = Mid(path, 1, InStrRev(path, "\") - 1)
10
    ShotDevicePath = path
End Function

' -----------------------------------------------------------------------------
' HasFlag (Function), source lines 757-759
' -----------------------------------------------------------------------------
Private Function HasFlag(ByVal value As Long, ByVal flag As Long) As Boolean
    HasFlag = (value And flag) = flag
End Function

' -----------------------------------------------------------------------------
' Configure (Sub), source lines 760-809
' -----------------------------------------------------------------------------
Public Sub Configure(PortName As String, hwnd As LongPtr)

    Dim deviceId As String
    Dim i As Long
    i = GetIndex(PortName)
    deviceId = ShotDevicePath(Interfaces(i).FullDevicePath)
    
    If Interfaces(i).isComPortAdapter Then
        If InStr(1, Interfaces(i).PortName, "COM", vbTextCompare) Then
            COMSettingsForm.PortName = RemoveTrailingColon(Trim(Interfaces(i).PortName))
            COMSettingsForm.AdvancedSettings = True
            COMSettingsForm.deviceId = deviceId
            COMSettingsForm.Show 1
            Exit Sub
        ElseIf InStr(1, Interfaces(i).PortName, "LPT", vbTextCompare) Then
            Dim retVal As Long
            retVal = ConfigurePort(vbNullString, hwnd, RemoveTrailingColon(Trim(Interfaces(i).PortName)) & ":")
            
            If retVal = 0 Then
                If Err.LastDllError = 1223 Then
                    'MsgBox "The dialog was canceled.", vbInformation, "Dialog closed"
                Else
                    MsgBox "The port is not available", vbInformation, "Error " & Err.LastDllError
                End If
            End If
        End If
    End If
    
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
    
    

End Sub

' -----------------------------------------------------------------------------
' RemoveTrailingColon (Function), source lines 810-818
' -----------------------------------------------------------------------------
Private Function RemoveTrailingColon(ByVal s As String) As String
    On Error GoTo 10
    If right(s, 1) = ":" Then
        RemoveTrailingColon = left(s, Len(s) - 1)
    Exit Function
    End If
10
    RemoveTrailingColon = s
End Function

' -----------------------------------------------------------------------------
' ExtractExparedId (Function), source lines 819-830
' -----------------------------------------------------------------------------
Private Function ExtractExparedId(devicePath As String) As String
    On Error GoTo 10
    Dim parts() As String
    parts = Split(devicePath, "#")
    
    If UBound(parts) > 1 Then
        ExtractExparedId = parts(2)
    End If
    
    Exit Function
10:
End Function

' -----------------------------------------------------------------------------
' ExtractVidPid (Function), source lines 831-842
' -----------------------------------------------------------------------------
Private Function ExtractVidPid(devicePath As String) As String
    On Error GoTo 10
    Dim parts() As String
    parts = Split(devicePath, "#")
    
    If UBound(parts) > 0 Then
        ExtractVidPid = parts(1)
    End If
    
    Exit Function
10:
End Function

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 843-848
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Dim q As Long
    For q = 0 To lCount - 1
        If Interfaces(q).Initialized Then DeInitialize (q)
    Next
End Sub
