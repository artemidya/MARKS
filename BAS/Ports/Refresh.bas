Public Sub Refresh()
    Class_Terminate 'DeInitialize
    Erase Ports()
    lCount = 0
    On Error GoTo 10
    Dim bufferSize As Long
    Dim buffer() As Byte
    Dim lPorts() As PORT_INFO_2
    Dim cPort As deviceInfo

    ' Determine number of ports
    EnumPorts vbNullChar, 2, ByVal 0&, 0&, bufferSize, lCount
    ' Create buffer and get port info
    ReDim buffer(bufferSize - 1)
    EnumPorts vbNullChar, 2, buffer(0), bufferSize, bufferSize, lCount
    ReDim lPorts(lCount - 1)

    MoveMemory lPorts(0), buffer(0), LenB(lPorts(0)) * lCount
    ' Get the name of each port
    Dim q As Long
    For q = 0 To lCount - 1
        With cPort
            .name = PtrToString(lPorts(q).pPortName)
            .Description = PtrToString(lPorts(q).pDescription)
            .MonitorName = PtrToString(lPorts(q).pMonitorName)
            .Active = False
           
            If InStr(1, UCase(.name), "COM") Then
                .Active = IsCOMPortAvailable(.name)
                AddToArray Ports(), cPort
            ElseIf InStr(1, UCase(.name), "LPT") Then
                .Active = True
                AddToArray Ports(), cPort
            End If
            
        End With
        
    Next
    
    If Not Not Ports() Then
        lCount = UBound(Ports) + 1
    Else
        lCount = 0
    End If
    Sort Ports()
10
End Sub
