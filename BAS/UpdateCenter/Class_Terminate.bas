Private Sub Class_Terminate()
    If Not socket1 Is Nothing Then
        socket1.Disconnect
    End If
End Sub
