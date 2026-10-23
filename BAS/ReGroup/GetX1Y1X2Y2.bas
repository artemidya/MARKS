Private Sub GetX1Y1X2Y2(ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, segment As VGCore.segment)
    With segment
        If .StartNode.PositionY < .EndNode.PositionY Then
            With .StartNode
                x1 = .PositionX
                y1 = .PositionY
            End With
            With .EndNode
                x2 = .PositionX
                y2 = .PositionY
            End With
        Else
            With .EndNode
                x1 = .PositionX
                y1 = .PositionY
            End With
            With .StartNode
                x2 = .PositionX
                y2 = .PositionY
            End With
        End If
    End With
End Sub
