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
