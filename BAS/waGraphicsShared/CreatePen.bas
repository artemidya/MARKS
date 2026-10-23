Public Function CreatePen(Optional ColorRGB As Long = 6579300, Optional Opacity As Long = 100, Optional Width As Single = 1) As waPen
    Set CreatePen = New waPen
    CreatePen.SetColor ColorRGB, Opacity, Width
End Function
