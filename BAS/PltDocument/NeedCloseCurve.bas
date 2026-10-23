Private Function NeedCloseCurve(ByRef segments() As segment, q As Long, startX As Long, startY As Long) As Boolean
    If q = UBound(segments()) Then
        If segments(q).P2.x = startX And -segments(q).P2.y = startY Then NeedCloseCurve = True
    Else
        If segments(q + 1).Type = InvisibleLine And segments(q).P2.x = startX And -segments(q).P2.y = startY Then NeedCloseCurve = True
    End If
End Function
