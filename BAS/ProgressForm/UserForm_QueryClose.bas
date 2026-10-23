Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not MbClosing Then
        Break = True
        Cancel = 1
    End If
End Sub
