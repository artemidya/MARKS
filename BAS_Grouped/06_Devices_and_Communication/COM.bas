' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: COM.bas
' Procedures: 12
' =============================================================================

' -----------------------------------------------------------------------------
' GetSystemMessage (Function), source lines 278-293
' -----------------------------------------------------------------------------
Public Function GetSystemMessage(lngErrorCode As Long) As String
Dim intPos As Integer
Dim strMessage As String, strMsgBuff As String * 256

    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngErrorCode, 0, strMsgBuff, 255, 0)

    intPos = InStr(1, strMsgBuff, vbNullChar)
    If intPos > 0 Then
        strMessage = Trim$(left(strMsgBuff, intPos - 1))
    Else
        strMessage = Trim$(strMsgBuff)
    End If
    
    GetSystemMessage = strMessage
    
End Function

' -----------------------------------------------------------------------------
' CommOpen (Function), source lines 309-419
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' SetCommError (Function), source lines 422-431
' -----------------------------------------------------------------------------
Private Function SetCommError(strFunction As String) As Long
    
    With udtCommError
        .lngErrorCode = Err.LastDllError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
        SetCommError = .lngErrorCode
    End With
    
End Function

' -----------------------------------------------------------------------------
' SetCommErrorEx (Function), source lines 433-450
' -----------------------------------------------------------------------------
Private Function SetCommErrorEx(strFunction As String, lngHnd As Long) As Long
Dim lngErrorFlags As Long
Dim udtCommStat As COMSTAT
    
    With udtCommError
        .lngErrorCode = GetLastError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
    
        Call ClearCommError(lngHnd, lngErrorFlags, udtCommStat)
    
        .strErrorMessage = .strErrorMessage & "  COMM Error Flags = " & _
                Hex$(lngErrorFlags)
        
        SetCommErrorEx = .lngErrorCode
    End With
    
End Function

' -----------------------------------------------------------------------------
' CommSet (Function), source lines 463-506
' -----------------------------------------------------------------------------
Public Function CommSet(intPortID As Integer, strSettings As String) As Long
    
Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (GetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    lngStatus = SetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommSet = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSet"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

' -----------------------------------------------------------------------------
' CommClose (Function), source lines 517-548
' -----------------------------------------------------------------------------
Public Function CommClose(intPortID As Integer) As Long
    
Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    If udtPorts(intPortID).blnPortOpen Then
        lngStatus = CloseHandle(udtPorts(intPortID).lngHandle)
    
        If lngStatus = 0 Then
            lngStatus = SetCommError("CommClose (CloseHandle)")
            GoTo Routine_Exit
        End If
    
        udtPorts(intPortID).blnPortOpen = False
    End If

    lngStatus = 0

Routine_Exit:
    CommClose = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommClose"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

' -----------------------------------------------------------------------------
' CommFlush (Function), source lines 559-587
' -----------------------------------------------------------------------------
Public Function CommFlush(intPortID As Integer) As Long
    
Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommFlush (PurgeComm)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommFlush = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommFlush"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

' -----------------------------------------------------------------------------
' CommRead (Function), source lines 600-686
' -----------------------------------------------------------------------------
Public Function CommRead(intPortID As Integer, strData As String, _
    lngSize As Long) As Long

Dim lngStatus As Long
Dim lngRdSize As Long, lngBytesRead As Long
Dim lngRdStatus As Long, strRdBuffer As String * 1024
Dim lngErrorFlags As Long, udtCommStat As COMSTAT
    
    On Error GoTo Routine_Error

    strData = ""
    lngBytesRead = 0
    DoEvents
    
    ' Clear any previous errors and get current status.
    lngStatus = ClearCommError(udtPorts(intPortID).lngHandle, lngErrorFlags, _
        udtCommStat)

    If lngStatus = 0 Then
        lngBytesRead = -1
        lngStatus = SetCommError("CommRead (ClearCommError)")
        GoTo Routine_Exit
    End If
        
    If udtCommStat.cbInQue > 0 Then
        If udtCommStat.cbInQue > lngSize Then
            lngRdSize = udtCommStat.cbInQue
        Else
            lngRdSize = lngSize
        End If
    Else
        lngRdSize = 0
    End If

    If lngRdSize Then
        lngRdStatus = ReadFile(udtPorts(intPortID).lngHandle, strRdBuffer, _
            lngRdSize, lngBytesRead, udtCommOverlap)

        If lngRdStatus = 0 Then
            lngStatus = GetLastError
            If lngStatus = ERROR_IO_PENDING Then
                ' Wait for read to complete.
                ' This function will timeout according to the
                ' COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
                ' Every time it times out, check for port errors.

                ' Loop until operation is complete.
                While GetOverlappedResult(udtPorts(intPortID).lngHandle, _
                    udtCommOverlap, lngBytesRead, True) = 0
                                    
                    lngStatus = GetLastError
                                        
                    If lngStatus <> ERROR_IO_INCOMPLETE Then
                        lngBytesRead = -1
                        lngStatus = SetCommErrorEx( _
                            "CommRead (GetOverlappedResult)", _
                            udtPorts(intPortID).lngHandle)
                        GoTo Routine_Exit
                    End If
                Wend
            Else
                ' Some other error occurred.
                lngBytesRead = -1
                lngStatus = SetCommErrorEx("CommRead (ReadFile)", _
                    udtPorts(intPortID).lngHandle)
                GoTo Routine_Exit
            
            End If
        End If
    
        strData = left(strRdBuffer, lngBytesRead)
    End If

Routine_Exit:
    CommRead = lngBytesRead
    Exit Function

Routine_Error:
    lngBytesRead = -1
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommRead"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

' -----------------------------------------------------------------------------
' CommWrite (Function), source lines 698-776
' -----------------------------------------------------------------------------
Public Function CommWrite(intPortID As Integer, strData As String) As Long
    
Dim i As Integer
Dim lngStatus As Long, lngSize As Long
Dim lngWrSize As Long, lngWrStatus As Long
    
    On Error GoTo Routine_Error
    
    ' Get the length of the data.
    lngSize = Len(strData)

    ' Output the data.
    lngWrStatus = WriteFile(udtPorts(intPortID).lngHandle, strData, lngSize, _
        lngWrSize, udtCommOverlap)

    ' Note that normally the following code will not execute because the driver
    ' caches write operations. Small I/O requests (up to several thousand bytes)
    ' will normally be accepted immediately and WriteFile will return true even
    ' though an overlapped operation was specified.
        
    DoEvents
    
    If lngWrStatus = 0 Then
        lngStatus = GetLastError
        If lngStatus = 0 Then
            GoTo Routine_Exit
        ElseIf lngStatus = ERROR_IO_PENDING Then
            ' We should wait for the completion of the write operation so we know
            ' if it worked or not.
            '
            ' This is only one way to do this. It might be beneficial to place the
            ' writing operation in a separate thread so that blocking on completion
            ' will not negatively affect the responsiveness of the UI.
            '
            ' If the write takes long enough to complete, this function will timeout
            ' according to the CommTimeOuts.WriteTotalTimeoutConstant variable.
            ' At that time we can check for errors and then wait some more.

            ' Loop until operation is complete.
            While GetOverlappedResult(udtPorts(intPortID).lngHandle, _
                udtCommOverlap, lngWrSize, True) = 0
                                
                lngStatus = GetLastError
                                    
                If lngStatus <> ERROR_IO_INCOMPLETE Then
                    lngStatus = SetCommErrorEx( _
                        "CommWrite (GetOverlappedResult)", _
                        udtPorts(intPortID).lngHandle)
                    GoTo Routine_Exit
                End If
            Wend
        Else
            ' Some other error occurred.
            lngWrSize = -1
                    
            lngStatus = SetCommErrorEx("CommWrite (WriteFile)", _
                udtPorts(intPortID).lngHandle)
            GoTo Routine_Exit
        
        End If
    End If
    
    For i = 1 To 10
        DoEvents
    Next
    
Routine_Exit:
    CommWrite = lngWrSize
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommWrite"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

' -----------------------------------------------------------------------------
' CommGetLine (Function), source lines 789-824
' -----------------------------------------------------------------------------
Public Function CommGetLine(intPortID As Integer, intLine As Integer, _
   blnState As Boolean) As Long
    
Dim lngStatus As Long
Dim lngComStatus As Long, lngModemStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommModemStatus(udtPorts(intPortID).lngHandle, lngModemStatus)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommReadCD (GetCommModemStatus)")
        GoTo Routine_Exit
    End If

    If (lngModemStatus And intLine) Then
        blnState = True
    Else
        blnState = False
    End If
        
    lngStatus = 0
        
Routine_Exit:
    CommGetLine = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommReadCD"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

' -----------------------------------------------------------------------------
' CommSetLine (Function), source lines 839-890
' -----------------------------------------------------------------------------
Public Function CommSetLine(intPortID As Integer, intLine As Integer, _
   blnState As Boolean) As Long
   
Dim lngStatus As Long
Dim lngNewState As Long
    
    On Error GoTo Routine_Error
    
    If intLine = LINE_BREAK Then
        If blnState Then
            lngNewState = SETBREAK
        Else
            lngNewState = CLRBREAK
        End If
    
    ElseIf intLine = LINE_DTR Then
        If blnState Then
            lngNewState = SETDTR
        Else
            lngNewState = CLRDTR
        End If
    
    ElseIf intLine = LINE_RTS Then
        If blnState Then
            lngNewState = SETRTS
        Else
            lngNewState = CLRRTS
        End If
    End If

    lngStatus = EscapeCommFunction(udtPorts(intPortID).lngHandle, lngNewState)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSetLine (EscapeCommFunction)")
        GoTo Routine_Exit
    End If

    lngStatus = 0
        
Routine_Exit:
    CommSetLine = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSetLine"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function

' -----------------------------------------------------------------------------
' CommGetError (Function), source lines 903-911
' -----------------------------------------------------------------------------
Public Function CommGetError(strMessage As String) As Long
    
    With udtCommError
        CommGetError = .lngErrorCode
        strMessage = "Error (" & CStr(.lngErrorCode) & "): " & .strFunction & _
            " - " & .strErrorMessage
    End With
    
End Function
