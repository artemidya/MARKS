Private Sub BlocksSocket1_OnDisconnect(client As BlocksSocket)
    If ActivationFormIsShow Then
        ActivationForm.Command1.Enabled = False
        ActivationForm.Text1.Enabled = False
    End If
End Sub
