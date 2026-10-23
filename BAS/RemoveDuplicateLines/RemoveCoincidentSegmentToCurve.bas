Private Sub RemoveCoincidentSegmentToCurve(TrimmingSegment As VGCore.segment, BreakSegments As VGCore.segments)

    With TrimmingSegment.boundingBox
        Dim left As Double
        Dim right As Double
        Dim top As Double
        Dim bottom As Double
        left = .left - Sensitivity
        right = .right + Sensitivity
        top = .top + Sensitivity
        bottom = .bottom - Sensitivity
    End With
     
    Dim RCSTS As Long
    Dim q As Long
    
    If TrimmingSegment.Type = cdrCurveSegment Then
        Dim PolylineCreated As Boolean
        Dim TrimmingPolyLine As VGCore.curve

        With BreakSegments
            q = 1
            Do While q <= .count
                If BoundingBoxCrossTest(BreakSegments(q).boundingBox, left, right, top, bottom) Then
                    If BreakSegments(q).Type = cdrCurveSegment Then
                        'Bezier-Bezier
                        RCSTS = RemoveCoincidentBezierSegmentToBezierSegment(TrimmingSegment, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    Else
                        'Bezier-Line
                        If Not PolylineCreated Then
                            Set TrimmingPolyLine = CreatePolylineCurve(TrimmingSegment)
                            PolylineCreated = True
                        End If
                        RCSTS = RemoveCoincidentBezierSegmentToLineSegment(TrimmingPolyLine, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    End If
                End If
                q = q + 1
            Loop
        End With
    Else
        With BreakSegments
            q = 1
            Do While q <= .count
                If BoundingBoxCrossTest(BreakSegments(q).boundingBox, left, right, top, bottom) Then
                    If BreakSegments(q).Type = cdrCurveSegment Then
                        'Line-Bezier
                        RCSTS = RemoveCoincidentLineSegmentToBezierSegment(TrimmingSegment, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    Else
                        'Line-Line
                        RCSTS = RemoveCoincidentLineSegmentToLineSegment(TrimmingSegment, BreakSegments(q))
                        q = q + RCSTS
                        If RCSTS > 0 Then Exit Sub
                    End If
                End If
                q = q + 1
            Loop
        End With
    End If
End Sub
