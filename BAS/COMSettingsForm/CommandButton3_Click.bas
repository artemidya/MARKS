Private Sub CommandButton3_Click()
    On Error GoTo 10
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
10
End Sub
