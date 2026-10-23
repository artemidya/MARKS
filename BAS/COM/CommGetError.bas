Public Function CommGetError(strMessage As String) As Long
    
    With udtCommError
        CommGetError = .lngErrorCode
        strMessage = "Error (" & CStr(.lngErrorCode) & "): " & .strFunction & _
            " - " & .strErrorMessage
    End With
    
End Function
