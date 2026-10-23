Private Function Distance(P1 As POINTL, P2 As POINTL) As Long
    
    Static x1 As Long, y1 As Long
    Static x2 As Long, y2 As Long
    
    With P1
        x1 = .x
        y1 = .y
    End With
    With P2
        x2 = .x
        y2 = .y
    End With
    
    Static a As Double, b As Double
    a = (y2 - y1)
    b = (x2 - x1)
    Distance = Sqr(a * a + b * b)
End Function
