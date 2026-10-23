Public Function CommOpen(intPortID As Integer, strPort As String, ComPort As String) As Long
    Dim strSettings As String
    Dim RegData1 As New RegData
'    strSettings = "baud=9600 parity=N data=8 stop=1"
    strSettings = RegData1.ComSettingsStr(ComPort)

    Dim lngStatus       As Long
    Dim udtCommTimeOuts As COMMTIMEOUTS

    On Error GoTo Routine_Error
    
    ' See if port already in use.
    If udtPorts(intPortID).blnPortOpen Then
        lngStatus = -1
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommOpen"
            .strErrorMessage = "Port in use."
        End With
        
        GoTo Routine_Exit
    End If

    ' Open serial port.
    udtPorts(intPortID).lngHandle = CreateFile(strPort, GENERIC_READ Or _
        GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If udtPorts(intPortID).lngHandle = -1 Then
        lngStatus = SetCommError("CommOpen (CreateFile)")
        GoTo Routine_Exit
    End If

    udtPorts(intPortID).blnPortOpen = True

    ' Setup device buffers (1K each).
    lngStatus = SetupComm(udtPorts(intPortID).lngHandle, 1024, 1024)
    
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetupComm)")
        GoTo Routine_Exit
    End If

    ' Purge buffers.
    lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (PurgeComm)")
        GoTo Routine_Exit
    End If

    ' Set serial port timeouts.
    With udtCommTimeOuts
        .ReadIntervalTimeout = -1
        .ReadTotalTimeoutMultiplier = 0
        .ReadTotalTimeoutConstant = 1000
        .WriteTotalTimeoutMultiplier = 0
        .WriteTotalTimeoutMultiplier = 1000
    End With

    lngStatus = SetCommTimeouts(udtPorts(intPortID).lngHandle, udtCommTimeOuts)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommTimeouts)")
        GoTo Routine_Exit
    End If

    ' Get the current state (DCB).
    lngStatus = GetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (GetCommState)")
        GoTo Routine_Exit
    End If



'GetCommState udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB

    ' Modify the DCB to reflect the desired settings.
    lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    ' Set the new state.
    lngStatus = SetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommOpen = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommOpen"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
