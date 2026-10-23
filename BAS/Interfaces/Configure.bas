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
        ' ќжидание завершени€ внешнего процесса (модальное поведение)
        WaitForSingleObject pi.hProcess, -1
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
    End If
    
    

End Sub
