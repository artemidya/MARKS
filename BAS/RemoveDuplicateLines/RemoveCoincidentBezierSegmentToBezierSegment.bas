Private Function RemoveCoincidentBezierSegmentToBezierSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    'Bezier-Bezier
    Dim OldThresholdAngle As Double
    Dim OldSensitivity As Double
    OldThresholdAngle = ThresholdAngle
    OldSensitivity = Sensitivity
    
    Sensitivity = Sensitivity * 1.15
    ThresholdAngle = ThresholdAngle * 5
    
    Dim nc As Long
    Dim mSegments As VGCore.segments
    Set mSegments = BreakSegment.Parent.segments
    nc = mSegments.count
    
    Dim minLength As Double
    Dim BreakPolyLine As VGCore.curve
    Dim TrimmingPolyLine As VGCore.curve
    
    Set BreakPolyLine = CreatePolylineCurve(BreakSegment, minLength)
    Dim c As Double
    If minLength > 0 Then
        c = TrimmingSegment.Length / (minLength * 100)
        If c > 500 Then c = 500
    Else
        c = 500
    End If
    Set TrimmingPolyLine = CreatePolylineCurve2(TrimmingSegment, CLng(c))
    
    Dim rs() As RemovableSegment

    RemoveCoincidentCurveToCurve TrimmingPolyLine.segments, BreakPolyLine
    DeleteLittleSubPathsFromPolyLine BreakPolyLine.Subpaths, minLength
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
    
    RemoveCoincidentBezierSegmentToBezierSegment = mSegments.count - nc
    
    ThresholdAngle = OldThresholdAngle
    Sensitivity = OldSensitivity
End Function
