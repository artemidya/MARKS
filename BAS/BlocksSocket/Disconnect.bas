Public Sub Disconnect()
'    If shutdown(lSocket, SD_BOTH) < 0 Then
'        oCloseSocket
'    End If
    
    shutdown lSocket, SD_SEND
    
    oCloseSocket
End Sub
