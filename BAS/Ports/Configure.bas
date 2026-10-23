Public Sub Configure(PortName As String, hwnd As LongPtr)
    
    If InStr(1, PortName, "COM", vbTextCompare) Then
        COMSettingsForm.PortName = RemoveTrailingColon(Trim(PortName))
        COMSettingsForm.AdvancedSettings = False
        COMSettingsForm.deviceId = ""
        COMSettingsForm.Show 1
        Exit Sub
    End If
    

    Dim retVal As Long
    retVal = ConfigurePort(vbNullString, hwnd, PortName)
    
    If retVal = 0 Then
        If Err.LastDllError = ERROR_CANCELLED Then
            'MsgBox "The dialog was canceled.", vbInformation, "Dialog closed"
        Else
            MsgBox "The port is not available", vbInformation, "Error " & Err.LastDllError
        End If
    End If
End Sub
