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
