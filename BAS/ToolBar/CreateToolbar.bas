Public Sub CreateToolbar()
    Select Case UCase(Application.name)
        Case "CORELDRAW"
            CreateToolBarForCorelDraw
        Case "CORELDESIGNER"
            MsgBox "Corel Designer???"
        Case Else
            CreateToolBarForCorelDraw
    End Select
End Sub
