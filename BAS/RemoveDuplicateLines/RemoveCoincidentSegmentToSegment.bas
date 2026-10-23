Private Function RemoveCoincidentSegmentToSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    If TrimmingSegment.Type = cdrLineSegment Then
        If BreakSegment.Type = cdrLineSegment Then
            'Line-Line
            RemoveCoincidentSegmentToSegment = RemoveCoincidentLineSegmentToLineSegment(TrimmingSegment, BreakSegment)
        ElseIf BreakSegment.Type = cdrCurveSegment Then
            'Line-Bezier
            RemoveCoincidentSegmentToSegment = RemoveCoincidentLineSegmentToBezierSegment(TrimmingSegment, BreakSegment)
        End If
    ElseIf TrimmingSegment.Type = cdrCurveSegment Then
        If BreakSegment.Type = cdrLineSegment Then
            'Bezier-Line
            Dim TrimmingPolyLine As VGCore.curve
            Set TrimmingPolyLine = CreatePolylineCurve(TrimmingSegment)
            RemoveCoincidentSegmentToSegment = RemoveCoincidentBezierSegmentToLineSegment(TrimmingPolyLine, BreakSegment)
        ElseIf BreakSegment.Type = cdrCurveSegment Then
            'Bezier-Bezier
            If Not RemoveCoincidentBezierSegmentToBezierSegmentQuickTest(TrimmingSegment, BreakSegment) Then
                RemoveCoincidentSegmentToSegment = RemoveCoincidentBezierSegmentToBezierSegment(TrimmingSegment, BreakSegment)
            End If
        End If
    End If
End Function
