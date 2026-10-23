Private Function ShiftSegment(segment As segment, ByVal Shift As Long) As segment
    Dim k As Double
    Dim dx As Long, dy As Long
    Dim sx As Long, sy As Long
    With segment
        dx = (.P2.x - .P1.x)
        dy = (.P2.y - .P1.y)
        If dx = 0 Then
            If dy < 0 Then sy = -Shift Else sy = Shift
        Else
            k = dy / dx
            If dx < 0 Then Shift = -Shift
            sx = Shift / Sqr(1 + k * k)
            sy = k * sx
        End If
    End With
    
    With ShiftSegment
        With .P1
            .x = segment.P1.x + sx
            .y = segment.P1.y + sy
        End With
        With .P2
            .x = segment.P2.x + sx
            .y = segment.P2.y + sy
        End With
    End With
End Function
