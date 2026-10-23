Private Function Distance(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Long
    Static a As Double, b As Double
    a = (y2 - y1)
    b = (x2 - x1) \ HorisontalDelitel
    Distance = Sqr(a * a + b * b)
End Function
