Private Function RemoveCoincidentLineSegmentToBezierSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    'Line-Bezier
    Dim OldThresholdAngle As Double
    OldThresholdAngle = ThresholdAngle
    ThresholdAngle = ThresholdAngle * 5
    
    Dim nc As Long
    Dim mSegments As VGCore.segments
    Set mSegments = BreakSegment.Parent.segments
    nc = mSegments.count
    
    Dim minLength As Double
    Dim BreakPolyLine As VGCore.curve
    Set BreakPolyLine = CreatePolylineCurve(BreakSegment, minLength)

    Dim rs() As RemovableSegment
    With BreakPolyLine
        RemoveCoincidentSegmentToCurve TrimmingSegment, .segments
        DeleteLittleSubPathsFromPolyLine .Subpaths, minLength
    End With
    With BreakSegment
        rs() = GetInvertSegments(BreakPolyLine.Subpaths, .StartNode.PositionX, .StartNode.PositionY, .EndNode.PositionX, .EndNode.PositionY)
    End With
    
    Dim q As Long
    Dim t As Double
    Dim NewNode As VGCore.Node
    
    If Not Not rs() Then
        For q = 0 To UBound(rs)
            With rs(q)
                'Start Cut Line
                t = GetParametrT(.startX, .startY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.AddNodeAt(t) '(this is node from next segment)
                    Set BreakSegment = NewNode.NextSegment
                End If
                
                'End Cut Line
                t = GetParametrT(.EndX, .EndY, BreakSegment)
                If t > 0 And t < 1 Then
                    Set NewNode = BreakSegment.BreakApartAt(t) '(this is node from next segment)
                    BreakSegment.EndNode.Delete
                    SetNextSegment NewNode, BreakSegment
                Else
                    With BreakSegment.EndNode
                        .BreakApart
                        .Delete
                        Exit For
                    End With
                End If
            End With
        Next
    End If

    RemoveCoincidentLineSegmentToBezierSegment = mSegments.count - nc
    ThresholdAngle = OldThresholdAngle
End Function
