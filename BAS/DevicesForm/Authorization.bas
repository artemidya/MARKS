Private Function Authorization(Optional Err As Boolean) As Boolean
    If Err = True Then lAuthorization = False
    If Not lAuthorization Then
        Dim pass As String
        pass = InputBox("Enter Moderator password", "Developer Mode", lPass)
        If Len(pass) > 4 Then
            lPass = pass
            'Send
            Dim buffer1 As New buffer
            With buffer1
                .InterPlatformNetConverter = True
                .size = Len(pass) * 2 + 4
                .bPut_Long Commands.Authorization
                .bPut_String pass
                BlocksSocket1.Send .bytes
            End With
            lAuthorization = True
        End If
    End If
    Authorization = lAuthorization
End Function
