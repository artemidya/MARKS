Private Function SetCommError(strFunction As String) As Long
    
    With udtCommError
        .lngErrorCode = Err.LastDllError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
        SetCommError = .lngErrorCode
    End With
    
End Function
