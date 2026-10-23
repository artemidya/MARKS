Public Sub ServerDOWN()
    'disconnect all
    If Not Clients Is Nothing Then
        Dim skt As BlocksSocket
        For Each skt In Clients
            skt.Disconnect
        Next
    End If
'    If Not Clients Is Nothing Then
'        MsgBox "SOME CLIENTS NOT IS CLOSED"
'    End If
    Set Clients = Nothing
    oCloseSocket
End Sub
