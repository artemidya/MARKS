Private Function AngelEndRev(path As waPath) As Double
    With path.LastSegment
        If .Type = CurveSegment Then
            AngelEndRev = Angel(.P4.x, .P4.y, .P3.x, .P3.y)
        Else
            AngelEndRev = Angel(.P2.x, .P2.y, .P1.x, .P1.y)
        End If
    End With
End Function
