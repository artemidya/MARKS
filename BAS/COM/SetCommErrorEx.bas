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
