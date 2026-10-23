Private Function CreatePolylineCurve(BezierSegment As VGCore.segment, Optional ByRef minLength As Double) As VGCore.curve
    With BezierSegment
        minLength = .Length
        Dim CountPolyLines As Long
        CountPolyLines = minLength * 0.2
        If CountPolyLines < 50 Then
            CountPolyLines = 50
        ElseIf CountPolyLines > 500 Then
            CountPolyLines = 500
        End If
        Set CreatePolylineCurve = .GetPolyline(CountPolyLines)
        minLength = minLength / CountPolyLines / 100
    End With
End Function
