Private Sub UserForm_Activate()
    On Error GoTo 10
    If FirstRun = False Then
        FirstRun = True
        FirstRunSub
        Timer1.Start 400
    End If
10
End Sub
