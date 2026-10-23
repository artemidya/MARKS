Private Sub RemoveCoincidentCurveToCurve(TrimmingSegments As VGCore.segments, BreakCurve As VGCore.curve)

    Dim BreakSegments As VGCore.segments
    Set BreakSegments = BreakCurve.segments

    With BreakCurve.boundingBox
        Dim left As Double
        Dim right As Double
        Dim top As Double
        Dim bottom As Double
        left = .left - Sensitivity
        right = .right + Sensitivity
        top = .top + Sensitivity
        bottom = .bottom - Sensitivity
    End With
    
    Dim q As Long
    q = 1
    With TrimmingSegments
        Do While q <= .count
            If BoundingBoxCrossTest(TrimmingSegments(q).boundingBox, left, right, top, bottom) Then
                RemoveCoincidentSegmentToCurve TrimmingSegments(q), BreakSegments
            End If
            q = q + 1
        Loop
    End With
End Sub
