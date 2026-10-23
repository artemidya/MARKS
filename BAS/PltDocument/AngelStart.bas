Private Function AngelStart(path As waPath) As Double
    With path.FirstSegment
        AngelStart = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
    End With
End Function
