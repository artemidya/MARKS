Private Sub Class_Terminate()
    If lSocket > 0 Then
        shutdown lSocket, SD_SEND
        oCloseSocket
    End If
    EndWinsock
End Sub
