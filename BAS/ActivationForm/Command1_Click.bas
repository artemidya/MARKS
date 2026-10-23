Private Sub Command1_Click()
    Dim buffer1 As New buffer
    buffer1.InterPlatformNetConverter = True
    buffer1.size = 44
    buffer1.bPut_Long Commands.ActivationToActCode
    Dim code As String * 20
    code = Text1.text
    buffer1.bPut_String code
    ParentForm.BlocksSocket1.Send buffer1.bytes
End Sub
