Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Label6.Font.Underline Then
        Label6.Font.Underline = False
        DoEvents
    End If
    Cursor.SetCursor2 Cursor.IDC_HAND
End Sub
