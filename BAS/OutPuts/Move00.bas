Public Sub Move00()
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(0, 0)
            Out StrToAscii(txt), 0, 0
        End If
    Else
        FileCoord.x = 0
        FileCoord.y = 0
    End If
End Sub
