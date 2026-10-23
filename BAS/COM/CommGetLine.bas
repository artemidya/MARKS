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
