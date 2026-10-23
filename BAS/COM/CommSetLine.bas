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
