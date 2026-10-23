Property Let Percent(p As Double)
    If p > 1 Then
        oPercent = 1
    ElseIf p < 0 Then
        oPercent = 0
    Else
        oPercent = p
    End If
    ProgressForm.PBarStrip.Width = Int(k * oPercent)
    ProgressForm.PBar.Caption = Int(oPercent * 100) & "%"
    Me.Repaint
End Property
