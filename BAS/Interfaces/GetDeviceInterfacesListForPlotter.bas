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
