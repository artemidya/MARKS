Private Function ParametrT(segment As VGCore.segment, SegmentForPoints As VGCore.segment) As ParametrT2X
    Static dx As Double
    Static dy As Double
    Static dPx1 As Double
    Static dPy1 As Double
    Static dPx2 As Double
    Static dPy2 As Double
    
    With segment
        dx = .EndNode.PositionX - .StartNode.PositionX
        dy = .EndNode.PositionY - .StartNode.PositionY
    End With

    With SegmentForPoints
        dPx1 = .StartNode.PositionX - segment.StartNode.PositionX
        dPy1 = .StartNode.PositionY - segment.StartNode.PositionY
        dPx2 = .EndNode.PositionX - segment.StartNode.PositionX
        dPy2 = .EndNode.PositionY - segment.StartNode.PositionY
    End With

    With ParametrT
        If Abs(dx) > Abs(dy) Then
            .t1 = dPx1 / dx
            .t2 = dPx2 / dx
        ElseIf dy <> 0 Then
            .t1 = dPy1 / dy
            .t2 = dPy2 / dy
        End If
    End With
End Function
