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
