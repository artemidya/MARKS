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
