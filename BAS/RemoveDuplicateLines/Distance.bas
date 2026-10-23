Private Function Distance(segment As VGCore.segment, SegmentForPoints As VGCore.segment) As Distance2X
    Static dx As Double
    Static dy As Double
    Static x1 As Double
    Static y1 As Double
    Static x2 As Double
    Static y2 As Double
    Static l As Double
    Static px1 As Double
    Static py1 As Double
    Static px2 As Double
    Static py2 As Double
    Static tmp As Double
    
    With SegmentForPoints
        With .StartNode
            px1 = .PositionX
            py1 = .PositionY
        End With
        With .EndNode
            px2 = .PositionX
            py2 = .PositionY
        End With
    End With

    With segment
        With .StartNode
            x1 = .PositionX
            y1 = .PositionY
        End With
        With .EndNode
            x2 = .PositionX
            y2 = .PositionY
        End With
    End With
    dx = x2 - x1
    dy = y2 - y1
    l = Sqr(dy * dy + dx * dx)
    If l > 0 Then
        With Distance
            tmp = x2 * y1 - y2 * x1
            .Distance1 = Abs(dy * px1 - dx * py1 + tmp) / l
            .Distance2 = Abs(dy * px2 - dx * py2 + tmp) / l
        End With
    Else
        With Distance
            .Distance1 = Sqr((py1 - y1) * (py1 - y1) + (px1 - x1) * (px1 - x1))
            .Distance2 = Sqr((py2 - y1) * (py2 - y1) + (px2 - x1) * (px2 - x1))
        End With
    End If
End Function
