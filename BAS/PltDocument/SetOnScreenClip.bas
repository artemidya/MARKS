Private Sub SetOnScreenClip(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)
    If x < 1 Then
        Width = Width - (1 - x)
        x = 1
    End If
    If y < 1 Then
        Height = Height - (1 - y)
        y = 1
    End If

    If x + Width > ScreenWidth Then Width = ScreenWidth - x
    If y + Height > ScreenHeight Then Height = ScreenHeight - y
    myGraphics.SetClip x, y, Width, Height
End Sub
