Private Function AngelEnd(path As waPath) As Double
    With path.LastSegment
        If .Type = CurveSegment Then
            AngelEnd = Angel(.P3.x, .P3.y, .P4.x, .P4.y)
        Else
            AngelEnd = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
        End If
    End With
End Function
