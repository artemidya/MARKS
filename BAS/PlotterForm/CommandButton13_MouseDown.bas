Private Sub CommandButton13_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Dim l As Double:    l = 5
    MovingProcesing = True
    DoEvents
    Do While MovingProcesing
        OutPuts.MoveLeft l
        RefreshXY
        Sleep 140
        If l < 10 Then l = l + 5
        DoEvents
    Loop
End Sub
