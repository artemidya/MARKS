Public Sub MoveUpRight(Optional lengthMm As Double = 5)
    On Error GoTo 10
    Dim Length As Long
    Length = lengthMm * 40
    Dim dx As Long, dy As Long
    Dim x As Long, y As Long
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
        dy = -Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    Else
        dy = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale)
        dx = Length * Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale)
    End If
    
    x = Me.x + dx
    y = Me.y + dy
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition <> CenterP Then
        If x < 0 Then x = 0
        If y < 0 Then y = 0
    End If
    
    If OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text)) <> file Then
        If Me.Connected(PlotterForm.Combo1.ListIndex) Then
            Dim txt As String
            txt = MoveCursor(x, y, 30)
            Out StrToAscii(txt), x, y
        End If
    Else
        FileCoord.x = x
        FileCoord.y = y
    End If
10
End Sub
