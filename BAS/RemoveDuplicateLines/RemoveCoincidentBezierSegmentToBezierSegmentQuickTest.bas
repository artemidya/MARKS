Private Function RemoveCoincidentBezierSegmentToBezierSegmentQuickTest(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Boolean
    With TrimmingSegment
        If .StartNode.PositionX = BreakSegment.StartNode.PositionX Then
            If .StartNode.PositionY = BreakSegment.StartNode.PositionY Then
                If .EndNode.PositionX = BreakSegment.EndNode.PositionX Then
                    If .EndNode.PositionY = BreakSegment.EndNode.PositionY Then
                        If .StartingControlPointX = BreakSegment.StartingControlPointX Then
                            If .StartingControlPointY = BreakSegment.StartingControlPointY Then
                                If .EndingControlPointX = BreakSegment.EndingControlPointX Then
                                    If .EndingControlPointY = BreakSegment.EndingControlPointY Then
                                        RemoveCoincidentBezierSegmentToBezierSegmentQuickTest = True
                                        With BreakSegment.EndNode
                                            .BreakApart
                                            .Delete
                                        End With
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function
