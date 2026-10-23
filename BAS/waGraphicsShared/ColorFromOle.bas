Public Function ColorFromOle(ColorOle As Long) As Long
    Dim c As Long
    c = ColorOle And &HFFFFFF
    ColorFromOle = GetSysColor(c)
End Function
